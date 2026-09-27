"""Leitura das configurações da aplicação.

Ordem de prioridade:
1. ``.streamlit/secrets.toml`` (local) ou "Secrets" do Streamlit Community Cloud;
2. variáveis de ambiente (útil para testes locais rápidos);
3. valores padrão (modo local, arquivo ``controle_financeiro.xlsx`` na pasta do projeto).
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any

import streamlit as st

DEFAULT_LOCAL_FILE = "controle_financeiro.xlsx"


@dataclass(frozen=True)
class GraphSettings:
    """Parâmetros de acesso ao OneDrive via Microsoft Graph API."""

    auth: str  # "refresh_token" (OneDrive pessoal) ou "client_credentials" (OneDrive corporativo)
    client_id: str
    tenant_id: str
    file_path: str  # caminho do arquivo dentro do OneDrive, ex.: "Financas/controle_financeiro.xlsx"
    refresh_token: str = ""
    client_secret: str = ""
    user_id: str = ""  # e-mail/UPN do dono do OneDrive (apenas client_credentials)


@dataclass(frozen=True)
class Settings:
    """Configuração consolidada da aplicação."""

    mode: str  # "local" ou "graph"
    local_path: str
    graph: GraphSettings | None
    app_password: str


def _secrets_section(name: str) -> dict[str, Any]:
    """Lê uma seção do secrets.toml sem quebrar quando o arquivo não existe."""
    try:
        section = st.secrets.get(name, {})
        return dict(section) if section else {}
    except Exception:  # noqa: BLE001 - sem secrets.toml o Streamlit lança exceção
        return {}


def _secret_value(name: str, default: str = "") -> str:
    try:
        return str(st.secrets.get(name, default))
    except Exception:  # noqa: BLE001
        return default


def load_settings() -> Settings:
    """Monta o objeto ``Settings`` a partir de secrets/ambiente."""
    storage = _secrets_section("storage")
    graph_cfg = _secrets_section("graph")

    mode = str(storage.get("mode") or os.getenv("FINANCAS_MODE") or "local").strip().lower()
    local_path = str(
        storage.get("local_path") or os.getenv("FINANCAS_XLSX_PATH") or DEFAULT_LOCAL_FILE
    ).strip()

    graph: GraphSettings | None = None
    if mode == "graph":
        auth = str(graph_cfg.get("auth", "refresh_token")).strip().lower()
        graph = GraphSettings(
            auth=auth,
            client_id=str(graph_cfg.get("client_id", "")).strip(),
            # "consumers" = contas Microsoft pessoais (outlook/hotmail/live)
            tenant_id=str(
                graph_cfg.get("tenant_id") or ("consumers" if auth == "refresh_token" else "")
            ).strip(),
            file_path=str(graph_cfg.get("file_path", "Financas/controle_financeiro.xlsx")).strip(),
            refresh_token=str(graph_cfg.get("refresh_token", "")).strip(),
            client_secret=str(graph_cfg.get("client_secret", "")).strip(),
            user_id=str(graph_cfg.get("user_id", "")).strip(),
        )

    app_password = _secret_value("app_password", os.getenv("FINANCAS_APP_PASSWORD", ""))
    return Settings(mode=mode, local_path=local_path, graph=graph, app_password=app_password)
