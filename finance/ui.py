"""Componentes de interface (Streamlit).

As ações usam callbacks (``on_click``): o callback altera os dados e salva no Excel
ANTES do rerun, então a tela já é redesenhada com os valores atualizados.
"""

from __future__ import annotations

import hmac

import altair as alt
import pandas as pd
import streamlit as st

from finance import operations as ops
from finance import state
from finance.calculations import (
    brl,
    current_month,
    debts_for_month,
    month_label,
    month_summary,
    monthly_history,
    piggy_balance,
    piggy_movements,
    salary_for_month,
    shift_month,
)
from finance.config import Settings
from finance.repository import COLUMNS, REG_SALARIO
from finance.storage import StorageBackend

CATEGORIAS = ["", "Aluguel", "Cartão", "Empréstimo", "Financiamento", "Contas da casa",
              "Assinatura", "Educação", "Saúde", "Transporte", "Outros"]


# --------------------------------------------------------------------------- #
# Infraestrutura de tela
# --------------------------------------------------------------------------- #
def password_gate(settings: Settings) -> bool:
    """Senha opcional (``app_password`` nos Secrets). Sem senha configurada, libera."""
    if not settings.app_password or st.session_state.get("fin_authenticated"):
        return True
    st.title("🔒 Controle Financeiro")
    with st.form("login"):
        pwd = st.text_input("Senha de acesso", type="password")
        ok = st.form_submit_button("Entrar", type="primary")
    if ok:
        if hmac.compare_digest(pwd.encode(), settings.app_password.encode()):
            st.session_state["fin_authenticated"] = True
            st.rerun()
        st.error("Senha incorreta.")
    return False


def render_flash() -> None:
    msg = state.pop_flash()
    if not msg:
        return
    kind, text = msg
    {"success": st.success, "warning": st.warning, "error": st.error}.get(kind, st.info)(text)


def render_load_error(message: str, backend: StorageBackend) -> None:
    st.error(f"**Não foi possível ler a planilha.**\n\n{message}")
    st.caption(f"Origem configurada: {backend.description}")
    st.button("Tentar novamente", on_click=state.ensure_loaded, args=(backend, True), type="primary")


# --------------------------------------------------------------------------- #
# Barra lateral: usuário, mês e arquivo
# --------------------------------------------------------------------------- #
def month_options() -> list[str]:
    cur = current_month()
    return [shift_month(cur, d) for d in range(-24, 13)]


def sidebar(backend: StorageBackend, settings: Settings) -> str:
    """Desenha a barra lateral e retorna o mês de referência selecionado."""
    with st.sidebar:
        st.header("👤 Usuário")
        users = state.users()
        if users:
            if st.session_state.get(state.K_USER) not in users:
                st.session_state[state.K_USER] = users[0]
            st.selectbox("Usuário ativo", users, key=state.K_USER)
        else:
            st.info("Nenhum usuário ainda. Crie o primeiro abaixo.")

        with st.expander("➕ Novo usuário", expanded=not users):
            with st.form("form_new_user", clear_on_submit=True):
                st.text_input("Nome (vira o nome da aba)", key="new_user_name", max_chars=31)
                st.form_submit_button(
                    "Criar usuário",
                    on_click=lambda: state.create_user(backend, st.session_state.get("new_user_name", "")),
                )

        st.header("📅 Mês de referência")
        options = month_options()
        st.session_state.setdefault("fin_month", current_month())
        month = st.selectbox("Mês", options, key="fin_month", format_func=month_label)

        st.divider()
        st.caption("📁 " + backend.description)
        if settings.mode == "local":
            st.caption("Modo local: use o caminho da pasta sincronizada pelo OneDrive.")
        st.button(
            "🔄 Recarregar planilha",
            on_click=state.ensure_loaded, args=(backend, True),
            help="Relê o arquivo (use se você editou a planilha direto no Excel).",
            width="stretch",
        )
    return month


# --------------------------------------------------------------------------- #
# Resumo (métricas + gráfico)
# --------------------------------------------------------------------------- #
def render_summary(df: pd.DataFrame, month: str) -> None:
    s = month_summary(df, month)
    st.subheader(f"Resumo de {month_label(month)}")
    if s.salary is None:
        st.info("Salário deste mês ainda não informado — lance-o na aba **💵 Salário**.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("💵 Salário", brl(s.salary))
    c2.metric("🧾 Total de dívidas", brl(s.total_debts),
              delta=f"{brl(s.pending)} pendente" if s.pending else "tudo pago",
              delta_color="inverse" if s.pending else "normal")
    c3.metric("💰 Saldo do mês", brl(s.balance),
              delta=f"Previsto: {brl(s.projected_balance)}", delta_color="off",
              help="Salário − soma das dívidas PAGAS no mês. "
                   "'Previsto' = saldo se todas as dívidas do mês forem pagas.")
    c4.metric("🐷 Cofrinho", brl(s.piggy_bank),
              delta=f"{brl(s.piggy_month_net)} no mês" if s.piggy_month_net else None)

    chart_df = pd.DataFrame({
        "Indicador": ["Salário", "Dívidas pagas", "Dívidas pendentes", "Saldo do mês", "Cofrinho"],
        "Valor": [s.salary or 0.0, s.paid, s.pending, s.balance, s.piggy_bank],
        "Cor": ["#2E86AB", "#E4572E", "#F3A712", "#29BF12" if s.balance >= 0 else "#C1121F", "#8E44AD"],
    })
    chart = (
        alt.Chart(chart_df)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("Indicador:N", sort=None, title=None, axis=alt.Axis(labelAngle=0)),
            y=alt.Y("Valor:Q", title="R$"),
            color=alt.Color("Cor:N", scale=None),
            tooltip=[alt.Tooltip("Indicador:N"), alt.Tooltip("Valor:Q", format=",.2f")],
        )
        .properties(height=300)
    )
    st.altair_chart(chart, width="stretch")


# --------------------------------------------------------------------------- #
# Salário
# --------------------------------------------------------------------------- #
def render_salary_tab(backend: StorageBackend, user: str, df: pd.DataFrame, month: str) -> None:
    current = salary_for_month(df, month)
    salary_key = f"salary_value_{user}_{month}"  # chave por usuário/mês: o campo mostra o valor certo
    st.markdown(f"Salário de **{month_label(month)}**: {brl(current)}")

    def save_salary() -> None:
        value = st.session_state.get(salary_key)
        state.run_action(
            backend, user,
            lambda d: ops.set_salary(d, month, float(value or 0)),
            f"Salário de {month_label(month)} definido em {brl(value)}",
        )

    with st.form(f"form_salary_{user}_{month}"):
        st.number_input("Valor recebido no mês (R$)", min_value=0.0, step=100.0, format="%.2f",
                        value=float(current or 0.0), key=salary_key)
        st.form_submit_button("Salvar salário", type="primary", on_click=save_salary)

    history = df[df["Registro"] == REG_SALARIO].sort_values("Mes_Referencia", ascending=False)
    if not history.empty:
        st.markdown("**Histórico de salários**")
        st.dataframe(
            pd.DataFrame({"Mês": history["Mes_Referencia"].map(month_label),
                          "Salário": history["Valor"].map(brl)}),
            hide_index=True, width="stretch",
        )


# --------------------------------------------------------------------------- #
# Dívidas
# --------------------------------------------------------------------------- #
def render_new_debt_form(backend: StorageBackend, user: str, month: str) -> None:
    with st.expander("➕ Adicionar nova dívida / despesa"):
        # O tipo fica fora do form para mostrar/ocultar o campo de parcelas na hora
        kind = st.radio("Tipo", [ops.TIPO_RECORRENTE, ops.TIPO_PARCELADA], horizontal=True,
                        key="debt_kind",
                        help="Recorrente: mensal fixa, sem fim (ex.: aluguel). "
                             "Parcelada: tem número de parcelas definido.")

        def add() -> None:
            ss = st.session_state
            parcelada = ss.get("debt_kind") == ops.TIPO_PARCELADA
            state.run_action(
                backend, user,
                lambda d: ops.add_debt(
                    d, name=ss.get("debt_name", ""), kind=ss.get("debt_kind"),
                    installment=ss.get("debt_value"), start_month=month,
                    category=ss.get("debt_category", ""),
                    total_installments=ss.get("debt_total") if parcelada else None,
                    paid_installments=ss.get("debt_paid", 0) if parcelada else 0,
                ),
                f"Dívida '{ss.get('debt_name', '').strip()}' adicionada",
            )

        with st.form("form_new_debt", clear_on_submit=True):
            c1, c2 = st.columns([2, 1])
            c1.text_input("Nome da dívida *", key="debt_name", placeholder="Ex.: Aluguel, Notebook...")
            c2.selectbox("Categoria (opcional)", CATEGORIAS, key="debt_category",
                         format_func=lambda c: c or "—")
            c3, c4, c5 = st.columns(3)
            c3.number_input("Valor da parcela (R$) *", min_value=0.0, step=10.0, format="%.2f", key="debt_value")
            if kind == ops.TIPO_PARCELADA:
                c4.number_input("Total de parcelas *", min_value=1, step=1, value=12, key="debt_total")
                c5.number_input("Parcelas já pagas", min_value=0, step=1, value=0, key="debt_paid",
                                help="Quantas parcelas você já pagou antes de cadastrar aqui.")
            else:
                c4.caption("Recorrente: sem número de parcelas (repete todo mês).")
            st.caption(f"A dívida passa a contar a partir de {month_label(month)}.")
            st.form_submit_button("Adicionar dívida", type="primary", on_click=add)


def render_debts_tab(backend: StorageBackend, user: str, df: pd.DataFrame, month: str) -> None:
    render_new_debt_form(backend, user, month)
    debts = debts_for_month(df, month)
    st.markdown(f"#### Dívidas de {month_label(month)}")
    if debts.empty:
        st.info("Nenhuma dívida ativa neste mês.")
        return

    for _, debt in debts.iterrows():
        debt_id, name = debt["ID"], debt["Descricao"]
        with st.container(border=True):
            c1, c2, c3, c4 = st.columns([3, 2, 2, 2], vertical_alignment="center")
            category = f" · {debt['Categoria']}" if debt["Categoria"] else ""
            c1.markdown(f"**{name}**  \n{debt['Tipo']}{category}")
            c2.markdown(f"Parcela: **{brl(debt['Valor_mes'])}**  \n{debt['Progresso']}")
            if debt["Tipo"] == ops.TIPO_PARCELADA and not pd.isna(debt["Parcelas_Total"]):
                c2.progress(min(float(debt["Parcelas_Pagas"]) / float(debt["Parcelas_Total"]), 1.0))
            if debt["Pago_no_mes"]:
                c3.success("Pago ✅")
                c4.button(
                    "Desfazer", key=f"undo_{debt_id}_{month}", width="stretch",
                    on_click=state.run_action,
                    args=(backend, user, lambda d, i=debt_id: ops.undo_payment(d, i, month),
                          f"Pagamento de '{name}' desfeito"),
                )
            else:
                c3.warning("Pendente ⏳")
                c4.button(
                    "Marcar como paga", key=f"pay_{debt_id}_{month}", type="primary",
                    width="stretch", on_click=state.run_action,
                    args=(backend, user, lambda d, i=debt_id: ops.pay_debt(d, i, month),
                          f"Parcela de '{name}' paga"),
                )
            if debt["Ativa"] == ops.SIM:
                with c1.popover("⋯ Encerrar"):
                    st.caption("Encerrar remove a dívida dos próximos meses (histórico é mantido).")
                    st.button(
                        "Confirmar encerramento", key=f"close_{debt_id}",
                        on_click=state.run_action,
                        args=(backend, user, lambda d, i=debt_id: ops.close_debt(d, i),
                              f"Dívida '{name}' encerrada"),
                    )


# --------------------------------------------------------------------------- #
# Cofrinho
# --------------------------------------------------------------------------- #
def render_piggy_tab(backend: StorageBackend, user: str, df: pd.DataFrame, month: str) -> None:
    balance = piggy_balance(df)
    st.metric("🐷 Saldo do cofrinho", brl(balance))

    def move() -> None:
        ss = st.session_state
        kind, value = ss.get("piggy_kind"), ss.get("piggy_value")
        state.run_action(
            backend, user,
            lambda d: ops.move_piggy_bank(d, kind, float(value or 0), month,
                                          ss.get("piggy_desc", ""), current_balance=balance),
            f"{kind} de {brl(value)} no cofrinho",
        )

    with st.form("form_piggy", clear_on_submit=True):
        c1, c2 = st.columns(2)
        c1.radio("Movimentação", [ops.TIPO_DEPOSITO, ops.TIPO_RETIRADA], horizontal=True, key="piggy_kind")
        c2.number_input("Valor (R$)", min_value=0.0, step=50.0, format="%.2f", key="piggy_value")
        st.text_input("Descrição (opcional)", key="piggy_desc", placeholder="Ex.: reserva de emergência")
        st.form_submit_button("Registrar", type="primary", on_click=move)

    moves = piggy_movements(df)
    if moves.empty:
        st.caption("Nenhuma movimentação ainda.")
        return
    moves["Saldo acumulado"] = moves["Valor_sinal"].cumsum()
    line = (
        alt.Chart(moves)
        .mark_area(line=True, opacity=0.25, color="#8E44AD", interpolate="step-after")
        .encode(
            x=alt.X("Data:T", title=None),
            y=alt.Y("Saldo acumulado:Q", title="R$"),
            tooltip=[alt.Tooltip("Data:T", format="%d/%m/%Y %H:%M"),
                     alt.Tooltip("Saldo acumulado:Q", format=",.2f")],
        )
        .properties(height=220)
    )
    st.altair_chart(line, width="stretch")
    table = moves.sort_values("Data", ascending=False)
    st.dataframe(
        pd.DataFrame({
            "Data": table["Data"].dt.strftime("%d/%m/%Y %H:%M"),
            "Tipo": table["Tipo"],
            "Valor": table["Valor_sinal"].map(brl),
            "Descrição": table["Descricao"],
        }),
        hide_index=True, width="stretch",
    )


# --------------------------------------------------------------------------- #
# Histórico
# --------------------------------------------------------------------------- #
def render_history_tab(df: pd.DataFrame) -> None:
    hist = monthly_history(df)
    if hist.empty:
        st.info("Ainda não há histórico. Lance um salário ou pague uma parcela.")
    else:
        long = hist.melt(id_vars=["Mês", "mes"], value_vars=["Salário", "Pago", "Saldo"],
                         var_name="Série", value_name="Valor")
        chart = (
            alt.Chart(long)
            .mark_line(point=True)
            .encode(
                x=alt.X("Mês:N", sort=list(hist["Mês"]), title=None),
                y=alt.Y("Valor:Q", title="R$"),
                color=alt.Color("Série:N", scale=alt.Scale(
                    domain=["Salário", "Pago", "Saldo"], range=["#2E86AB", "#E4572E", "#29BF12"])),
                tooltip=["Mês", "Série", alt.Tooltip("Valor:Q", format=",.2f")],
            )
            .properties(height=300)
        )
        st.altair_chart(chart, width="stretch")
        st.dataframe(hist.drop(columns="mes").assign(
            **{c: hist[c].map(brl) for c in ["Salário", "Pago", "Saldo"]}),
            hide_index=True, width="stretch")

    with st.expander("🗂️ Dados brutos da aba (como estão no Excel)"):
        st.dataframe(df[COLUMNS], hide_index=True, width="stretch")
