"""Estado da sessão (cache em ``st.session_state``) e salvamento automático.

Fluxo:
1. Na primeira execução da sessão, a planilha é lida UMA vez e cada aba de usuário vira
   um DataFrame guardado em ``st.session_state``.
2. Cada ação (pagar parcela, novo salário, cofrinho...) altera o DataFrame do usuário e
   chama ``persist`` imediatamente -> a mudança vai para o Excel na hora.
3. ``persist`` grava SÓ a aba do usuário ativo. Antes de gravar ele compara a versão do
   arquivo; se alguém alterou o arquivo por fora, baixa a versão nova (para não apagar as
   outras abas) e só então substitui a aba do usuário.
4. Se a gravação falhar, a alteração é desfeita na memória e o erro é exibido.
"""

from __future__ import annotations

from typing import Callable

import pandas as pd
import streamlit as st

from finance import repository
from finance.calculations import current_month
from finance.config import Settings
from finance.operations import ValidationError, refresh_month_status
from finance.storage import ConflictError, StorageBackend, StorageError, build_backend

# Chaves usadas no session_state
K_LOADED = "fin_loaded"
K_FRAMES = "fin_frames"          # dict[str, DataFrame] -> um DataFrame por usuário
K_BYTES = "fin_workbook_bytes"   # último conteúdo conhecido do arquivo
K_VERSION = "fin_version"        # versão (mtime/eTag) do conteúdo acima
K_LOAD_ERROR = "fin_load_error"
K_USER = "fin_active_user"
K_FLASH = "fin_flash"
K_SOURCE = "fin_source"          # identifica a origem carregada (para recarregar se mudar)


@st.cache_resource(show_spinner=False)
def get_backend(mode: str, local_path: str, _graph, graph_key: str) -> StorageBackend:
    """Backend compartilhado entre reruns (mantém o token do Graph em cache).

    ``_graph`` começa com "_" para o Streamlit não tentar gerar hash dele; ``graph_key``
    diferencia as configurações.
    """
    return build_backend(mode, local_path, _graph)


def backend_from_settings(settings: Settings) -> StorageBackend:
    graph_key = f"{settings.graph.client_id}|{settings.graph.file_path}" if settings.graph else ""
    return get_backend(settings.mode, settings.local_path, settings.graph, graph_key)


# --------------------------------------------------------------------------- #
# Mensagens (exibidas no próximo rerun, pois ações rodam em callbacks)
# --------------------------------------------------------------------------- #
def flash(kind: str, message: str) -> None:
    st.session_state[K_FLASH] = (kind, message)


def pop_flash() -> tuple[str, str] | None:
    return st.session_state.pop(K_FLASH, None)


# --------------------------------------------------------------------------- #
# Leitura
# --------------------------------------------------------------------------- #
def ensure_loaded(backend: StorageBackend, force: bool = False) -> None:
    """Carrega a planilha para o session_state (somente na 1ª vez ou quando ``force``)."""
    if st.session_state.get(K_SOURCE) != backend.description:
        force = True  # mudou o arquivo configurado
    if st.session_state.get(K_LOADED) and not force:
        return
    st.session_state[K_LOADED] = True
    st.session_state[K_SOURCE] = backend.description
    st.session_state[K_LOAD_ERROR] = None
    try:
        data, version = backend.read()
        frames = repository.load_user_frames(data)
    except (StorageError, repository.SchemaError) as exc:
        st.session_state[K_LOAD_ERROR] = str(exc)
        st.session_state[K_FRAMES] = {}
        st.session_state[K_BYTES] = None
        st.session_state[K_VERSION] = None
        return
    st.session_state[K_FRAMES] = frames
    st.session_state[K_BYTES] = data
    st.session_state[K_VERSION] = version
    if st.session_state.get(K_USER) not in frames:
        st.session_state[K_USER] = next(iter(frames), None)


def load_error() -> str | None:
    return st.session_state.get(K_LOAD_ERROR)


def users() -> list[str]:
    return list(st.session_state.get(K_FRAMES, {}).keys())


def active_user() -> str | None:
    return st.session_state.get(K_USER)


def frame(user: str) -> pd.DataFrame:
    return st.session_state[K_FRAMES][user]


# --------------------------------------------------------------------------- #
# Escrita
# --------------------------------------------------------------------------- #
def persist(backend: StorageBackend, user: str) -> bool:
    """Grava a aba do usuário no Excel. Retorna True se houve alteração externa."""
    df = refresh_month_status(frame(user), current_month())
    st.session_state[K_FRAMES][user] = df

    data = st.session_state.get(K_BYTES)
    version = st.session_state.get(K_VERSION)
    external_change = False

    for attempt in range(2):  # 2ª tentativa só em caso de conflito de versão
        remote_version = backend.get_version()
        if remote_version != version:
            # Arquivo mudou fora da sessão (ou foi apagado): parte do conteúdo mais recente
            data, version = backend.read()
            external_change = remote_version is not None
        new_bytes = repository.write_user_sheet(data, user, df)
        try:
            new_version = backend.write(new_bytes, expected_version=version)
            break
        except ConflictError:
            if attempt == 1:
                raise
            version = "__forçar_releitura__"

    st.session_state[K_BYTES] = new_bytes
    st.session_state[K_VERSION] = new_version
    if external_change:
        # Atualiza as abas dos OUTROS usuários com o que veio do arquivo
        fresh = repository.load_user_frames(new_bytes)
        fresh[user] = df
        st.session_state[K_FRAMES] = fresh
    return external_change


def run_action(
    backend: StorageBackend,
    user: str,
    mutate: Callable[[pd.DataFrame], pd.DataFrame],
    success_message: str,
) -> None:
    """Executa uma alteração + salvamento automático, com rollback em caso de erro."""
    before = frame(user)
    try:
        st.session_state[K_FRAMES][user] = mutate(before.copy())
        external = persist(backend, user)
    except ValidationError as exc:
        st.session_state[K_FRAMES][user] = before
        flash("warning", str(exc))
        return
    except (StorageError, repository.SchemaError) as exc:
        st.session_state[K_FRAMES][user] = before
        flash("error", f"Não foi possível salvar na planilha. Nada foi alterado. Detalhe: {exc}")
        return
    except Exception as exc:  # noqa: BLE001 - erro inesperado não pode deixar dado inconsistente
        st.session_state[K_FRAMES][user] = before
        flash("error", f"Erro inesperado ao salvar: {exc}")
        return
    suffix = " (a planilha tinha alterações externas; as outras abas foram preservadas)" if external else ""
    flash("success", f"{success_message} — salvo na planilha.{suffix}")


def create_user(backend: StorageBackend, name: str) -> None:
    """Cria a aba do novo usuário (com os cabeçalhos) e já salva no arquivo."""
    try:
        name = repository.validate_user_name(name, users())
    except ValueError as exc:
        flash("warning", str(exc))
        return
    st.session_state[K_FRAMES][name] = repository.empty_frame()
    try:
        persist(backend, name)
    except (StorageError, repository.SchemaError) as exc:
        st.session_state[K_FRAMES].pop(name, None)
        flash("error", f"Não foi possível criar a aba '{name}' na planilha: {exc}")
        return
    st.session_state[K_USER] = name
    flash("success", f"Usuário '{name}' criado (nova aba na planilha).")
