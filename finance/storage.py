"""Camada de armazenamento: lê e grava os BYTES do arquivo .xlsx.

Duas implementações com a mesma interface:

* ``LocalBackend``  -> arquivo em uma pasta local sincronizada pelo OneDrive
                       (ideal para testar localmente; não funciona no Streamlit Cloud,
                       pois o servidor não tem acesso ao seu computador).
* ``GraphBackend``  -> arquivo direto no OneDrive via Microsoft Graph API + MSAL
                       (é o modo usado no deploy no Streamlit Community Cloud).

Ambas trabalham com uma "versão" do arquivo (mtime local ou eTag do OneDrive).
A versão permite detectar se o arquivo foi alterado por outra pessoa/processo
entre a leitura e a gravação (controle de concorrência otimista).
"""

from __future__ import annotations

import os
import tempfile
import time
from pathlib import Path
from typing import Protocol
from urllib.parse import quote

import msal
import requests

from finance.config import GraphSettings

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
GRAPH_SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024  # upload simples do Graph aceita até 4 MB
HTTP_TIMEOUT = 30  # segundos


class StorageError(Exception):
    """Erro de leitura/escrita com mensagem amigável para o usuário."""


class ConflictError(StorageError):
    """O arquivo mudou desde a última leitura (outra pessoa/processo gravou)."""


class StorageBackend(Protocol):
    """Interface comum dos backends."""

    description: str

    def get_version(self) -> str | None:
        """Versão atual do arquivo (``None`` se não existir)."""

    def read(self) -> tuple[bytes | None, str | None]:
        """Retorna ``(conteúdo, versão)``; ``(None, None)`` se o arquivo não existir."""

    def write(self, data: bytes, expected_version: str | None) -> str:
        """Grava o conteúdo e retorna a nova versão."""


# --------------------------------------------------------------------------- #
# Backend local (pasta sincronizada pelo OneDrive)
# --------------------------------------------------------------------------- #
class LocalBackend:
    def __init__(self, path: str) -> None:
        self.path = Path(path).expanduser()
        self.description = f"Arquivo local: {self.path}"

    def get_version(self) -> str | None:
        try:
            return str(self.path.stat().st_mtime_ns)
        except FileNotFoundError:
            return None
        except OSError as exc:
            raise StorageError(f"Não foi possível acessar '{self.path}': {exc}") from exc

    def read(self) -> tuple[bytes | None, str | None]:
        try:
            version = self.get_version()
            if version is None:
                return None, None
            return self.path.read_bytes(), version
        except PermissionError as exc:
            raise StorageError(
                f"Sem permissão para ler '{self.path}'. Se o arquivo estiver aberto no Excel, "
                "feche-o e clique em 'Recarregar planilha'."
            ) from exc
        except OSError as exc:
            raise StorageError(f"Erro ao ler '{self.path}': {exc}") from exc

    def write(self, data: bytes, expected_version: str | None) -> str:
        current = self.get_version()
        if expected_version is not None and current is not None and current != expected_version:
            raise ConflictError("O arquivo foi alterado fora da aplicação desde a última leitura.")
        tmp_name: str | None = None
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            # Grava em arquivo temporário e substitui de forma atômica, evitando
            # que o OneDrive sincronize um arquivo pela metade.
            fd, tmp_name = tempfile.mkstemp(dir=self.path.parent, suffix=".tmp")
            with os.fdopen(fd, "wb") as tmp:
                tmp.write(data)
            os.replace(tmp_name, self.path)
        except PermissionError as exc:
            raise StorageError(
                f"Não foi possível salvar em '{self.path}'. O arquivo provavelmente está aberto "
                "no Excel (ele bloqueia a gravação). Feche o Excel e tente novamente."
            ) from exc
        except OSError as exc:
            raise StorageError(f"Erro ao salvar '{self.path}': {exc}") from exc
        finally:
            if tmp_name and os.path.exists(tmp_name):  # sobra de gravação com erro
                os.remove(tmp_name)
        return self.get_version() or ""


# --------------------------------------------------------------------------- #
# Backend OneDrive via Microsoft Graph API
# --------------------------------------------------------------------------- #
class GraphBackend:
    """Acessa o arquivo no OneDrive usando a Microsoft Graph API.

    Modos de autenticação (``[graph].auth`` no secrets.toml):

    * ``refresh_token``      -> OneDrive PESSOAL (conta Microsoft/Outlook). Usa um refresh
                                token gerado uma única vez pelo script ``get_token.py``.
    * ``client_credentials`` -> OneDrive CORPORATIVO (Microsoft 365). Usa client secret de um
                                app registrado no Entra ID com permissão de aplicativo
                                ``Files.ReadWrite.All``.
    """

    DELEGATED_SCOPES = ["Files.ReadWrite"]
    APP_SCOPES = ["https://graph.microsoft.com/.default"]

    def __init__(self, settings: GraphSettings) -> None:
        self._validate(settings)
        self.s = settings
        if settings.auth == "client_credentials":
            drive = f"{GRAPH_ROOT}/users/{quote(settings.user_id)}/drive"
        else:
            drive = f"{GRAPH_ROOT}/me/drive"
        path = quote(settings.file_path.strip("/"))
        self.meta_url = f"{drive}/root:/{path}"
        self.content_url = f"{drive}/root:/{path}:/content"
        self.description = f"OneDrive (Graph API): {settings.file_path}"
        self._token: str | None = None
        self._token_expires = 0.0
        self._refresh_token = settings.refresh_token  # o Graph pode devolver um novo a cada uso

    @staticmethod
    def _validate(s: GraphSettings) -> None:
        missing = [n for n in ("client_id", "file_path", "tenant_id") if not getattr(s, n)]
        if s.auth == "refresh_token" and not s.refresh_token:
            missing.append("refresh_token")
        elif s.auth == "client_credentials":
            missing += [n for n in ("client_secret", "user_id") if not getattr(s, n)]
        elif s.auth not in ("refresh_token", "client_credentials"):
            raise StorageError("[graph].auth deve ser 'refresh_token' ou 'client_credentials'.")
        if missing:
            raise StorageError(
                "Configuração do OneDrive incompleta no secrets.toml. Faltando em [graph]: "
                + ", ".join(missing)
            )

    # ----- autenticação ------------------------------------------------------
    def _access_token(self) -> str:
        if self._token and time.time() < self._token_expires - 60:
            return self._token
        authority = f"https://login.microsoftonline.com/{self.s.tenant_id}"
        try:
            if self.s.auth == "client_credentials":
                app = msal.ConfidentialClientApplication(
                    self.s.client_id, authority=authority, client_credential=self.s.client_secret
                )
                result = app.acquire_token_for_client(scopes=self.APP_SCOPES)
            else:
                app = msal.PublicClientApplication(self.s.client_id, authority=authority)
                result = app.acquire_token_by_refresh_token(
                    self._refresh_token, scopes=self.DELEGATED_SCOPES
                )
        except (requests.RequestException, ValueError) as exc:
            raise StorageError(f"Não foi possível conectar ao login da Microsoft: {exc}") from exc

        if "access_token" not in result:
            detail = result.get("error_description") or result.get("error") or "erro desconhecido"
            hint = (
                " O refresh token pode ter expirado: gere outro com `python get_token.py` "
                "e atualize os Secrets."
                if self.s.auth == "refresh_token"
                else ""
            )
            raise StorageError(f"Falha na autenticação com a Microsoft: {detail}.{hint}")

        self._refresh_token = result.get("refresh_token", self._refresh_token)
        self._token = result["access_token"]
        self._token_expires = time.time() + int(result.get("expires_in", 3600))
        return self._token

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        headers = kwargs.pop("headers", {})
        headers["Authorization"] = f"Bearer {self._access_token()}"
        try:
            return requests.request(method, url, headers=headers, timeout=HTTP_TIMEOUT, **kwargs)
        except requests.RequestException as exc:
            raise StorageError(f"Falha de conexão com o OneDrive: {exc}") from exc

    @staticmethod
    def _raise_for_status(resp: requests.Response, action: str) -> None:
        if resp.ok:
            return
        messages = {
            401: "credenciais inválidas ou expiradas",
            403: "sem permissão para acessar o arquivo (verifique as permissões do app no Entra ID)",
            404: "arquivo ou pasta não encontrado no OneDrive",
            412: "o arquivo foi alterado por outra pessoa",
            423: "o arquivo está bloqueado (aberto para edição em outro lugar)",
            429: "muitas requisições ao OneDrive; aguarde alguns segundos",
        }
        reason = messages.get(resp.status_code, f"HTTP {resp.status_code}")
        try:
            detail = resp.json().get("error", {}).get("message", "")
        except ValueError:
            detail = ""
        exc_cls = ConflictError if resp.status_code == 412 else StorageError
        raise exc_cls(f"Erro ao {action} no OneDrive: {reason}. {detail}".strip())

    # ----- interface ---------------------------------------------------------
    def _metadata(self) -> dict | None:
        resp = self._request("GET", self.meta_url)
        if resp.status_code == 404:
            return None
        self._raise_for_status(resp, "consultar o arquivo")
        return resp.json()

    def get_version(self) -> str | None:
        meta = self._metadata()
        return meta.get("eTag") if meta else None

    def read(self) -> tuple[bytes | None, str | None]:
        meta = self._metadata()
        if meta is None:
            return None, None
        download_url = meta.get("@microsoft.graph.downloadUrl")
        try:
            if download_url:  # URL pré-autenticada, não precisa de token
                resp = requests.get(download_url, timeout=HTTP_TIMEOUT)
            else:
                resp = self._request("GET", self.content_url)
        except requests.RequestException as exc:
            raise StorageError(f"Falha ao baixar o arquivo do OneDrive: {exc}") from exc
        self._raise_for_status(resp, "baixar o arquivo")
        return resp.content, meta.get("eTag")

    def write(self, data: bytes, expected_version: str | None) -> str:
        if len(data) > GRAPH_SIMPLE_UPLOAD_LIMIT:
            raise StorageError("A planilha passou de 4 MB, limite do upload simples do Graph.")
        headers = {
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
        if expected_version:
            headers["If-Match"] = expected_version  # evita sobrescrever alteração alheia
        resp = self._request("PUT", self.content_url, data=data, headers=headers)
        self._raise_for_status(resp, "salvar o arquivo")
        return resp.json().get("eTag", "")


def build_backend(mode: str, local_path: str, graph: GraphSettings | None) -> StorageBackend:
    """Fábrica de backends a partir da configuração."""
    if mode == "local":
        return LocalBackend(local_path)
    if mode == "graph":
        if graph is None:
            raise StorageError("Modo 'graph' selecionado, mas a seção [graph] não foi configurada.")
        return GraphBackend(graph)
    raise StorageError(f"Modo de armazenamento desconhecido: '{mode}'. Use 'local' ou 'graph'.")
