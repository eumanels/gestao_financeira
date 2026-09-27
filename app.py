"""Controle Financeiro Pessoal — Streamlit + Excel (OneDrive).

Executar localmente:
    pip install -r requirements.txt
    streamlit run app.py

Configuração em .streamlit/secrets.toml (veja secrets.toml.example e o README).
"""

from __future__ import annotations

import streamlit as st

from finance import state, ui
from finance.config import load_settings
from finance.storage import StorageError

st.set_page_config(page_title="Controle Financeiro", page_icon="💰", layout="wide")


def main() -> None:
    settings = load_settings()

    # 1) Senha opcional (recomendado no deploy público)
    if not ui.password_gate(settings):
        return

    # 2) Backend de armazenamento (pasta local sincronizada ou OneDrive via Graph)
    try:
        backend = state.backend_from_settings(settings)
    except StorageError as exc:
        st.error(f"**Configuração de armazenamento inválida.**\n\n{exc}")
        st.stop()

    # 3) Carrega a planilha UMA vez por sessão (depois tudo vem do session_state)
    with st.spinner("Carregando planilha..."):
        state.ensure_loaded(backend)

    st.title("💰 Controle Financeiro Pessoal")
    ui.render_flash()

    error = state.load_error()
    if error:
        ui.render_load_error(error, backend)
        st.stop()

    month = ui.sidebar(backend, settings)
    user = state.active_user()
    if not user:
        st.info("👈 Crie o primeiro usuário na barra lateral para começar. "
                "Cada usuário vira uma aba na planilha.")
        st.stop()

    df = state.frame(user)
    st.caption(f"Usuário: **{user}**")
    ui.render_summary(df, month)

    tab_debts, tab_salary, tab_piggy, tab_history = st.tabs(
        ["🧾 Dívidas", "💵 Salário", "🐷 Cofrinho", "📈 Histórico"]
    )
    with tab_debts:
        ui.render_debts_tab(backend, user, df, month)
    with tab_salary:
        ui.render_salary_tab(backend, user, df, month)
    with tab_piggy:
        ui.render_piggy_tab(backend, user, df, month)
    with tab_history:
        ui.render_history_tab(df)


main()
