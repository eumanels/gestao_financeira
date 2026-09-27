"""Cálculos de saldo e resumos (somente leitura, funções puras)."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date

import pandas as pd

from finance.operations import SIM, TIPO_DEPOSITO, TIPO_PARCELADA
from finance.repository import REG_COFRINHO, REG_DIVIDA, REG_PAGAMENTO, REG_SALARIO


@dataclass(frozen=True)
class MonthSummary:
    month: str
    salary: float | None      # None = salário do mês ainda não informado
    total_debts: float        # soma das parcelas que vencem no mês (pagas + pendentes)
    paid: float               # soma das parcelas pagas no mês
    pending: float            # total_debts - paid
    balance: float            # SALDO DO MÊS = salário - dívidas pagas no mês
    projected_balance: float  # salário - todas as dívidas do mês (se pagar tudo)
    piggy_bank: float         # saldo total do cofrinho
    piggy_month_net: float    # depósitos - retiradas no mês


# --------------------------------------------------------------------------- #
# Utilidades de mês
# --------------------------------------------------------------------------- #
def current_month(today: date | None = None) -> str:
    today = today or date.today()
    return f"{today.year:04d}-{today.month:02d}"


def shift_month(month: str, delta: int) -> str:
    year, mon = map(int, month.split("-"))
    total = year * 12 + (mon - 1) + delta
    return f"{total // 12:04d}-{total % 12 + 1:02d}"


def month_label(month: str) -> str:
    """'2026-09' -> '09/2026'."""
    year, mon = month.split("-")
    return f"{mon}/{year}"


def brl(value: float | None) -> str:
    """Formata em Real: 1234.5 -> 'R$ 1.234,50'."""
    if value is None or pd.isna(value):
        return "—"
    text = f"{abs(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"-R$ {text}" if value < 0 else f"R$ {text}"


# --------------------------------------------------------------------------- #
# Consultas
# --------------------------------------------------------------------------- #
def salary_for_month(df: pd.DataFrame, month: str) -> float | None:
    rows = df[(df["Registro"] == REG_SALARIO) & (df["Mes_Referencia"] == month)]
    return None if rows.empty else float(rows["Valor"].iloc[0])


def payments_for_month(df: pd.DataFrame, month: str) -> pd.DataFrame:
    return df[(df["Registro"] == REG_PAGAMENTO) & (df["Mes_Referencia"] == month)]


def debts_for_month(df: pd.DataFrame, month: str) -> pd.DataFrame:
    """Dívidas que compõem o mês: ativas (iniciadas até o mês) + as pagas no mês.

    Colunas extras: ``Pago_no_mes`` (bool), ``Valor_mes`` (valor pago ou valor da parcela)
    e ``Progresso`` (texto 'pagas/total').
    """
    debts = df[df["Registro"] == REG_DIVIDA].copy()
    payments = payments_for_month(df, month)
    paid_values = payments.groupby("ID_Divida")["Valor"].sum()

    started = (debts["Mes_Referencia"] == "") | (debts["Mes_Referencia"] <= month)
    paid_in_month = debts["ID"].isin(paid_values.index)
    debts = debts[(started & (debts["Ativa"] == SIM)) | paid_in_month].copy()

    debts["Pago_no_mes"] = debts["ID"].isin(paid_values.index)
    debts["Valor_mes"] = debts.apply(
        lambda r: float(paid_values.get(r["ID"], r["Valor"])), axis=1
    ) if not debts.empty else pd.Series(dtype=float)

    def progress(r) -> str:
        paid = int(0 if pd.isna(r["Parcelas_Pagas"]) else r["Parcelas_Pagas"])
        if r["Tipo"] == TIPO_PARCELADA and not pd.isna(r["Parcelas_Total"]):
            return f"{paid}/{int(r['Parcelas_Total'])}"
        return f"Recorrente · {paid} paga(s)"

    debts["Progresso"] = debts.apply(progress, axis=1) if not debts.empty else pd.Series(dtype=object)
    return debts.sort_values(["Pago_no_mes", "Descricao"]).reset_index(drop=True)


def piggy_movements(df: pd.DataFrame) -> pd.DataFrame:
    """Movimentações do cofrinho com valor com sinal (+ depósito, − retirada)."""
    moves = df[df["Registro"] == REG_COFRINHO].copy()
    moves["Valor_sinal"] = moves.apply(
        lambda r: r["Valor"] if r["Tipo"] == TIPO_DEPOSITO else -r["Valor"], axis=1
    ) if not moves.empty else pd.Series(dtype=float)
    return moves.sort_values("Data").reset_index(drop=True)


def piggy_balance(df: pd.DataFrame) -> float:
    moves = piggy_movements(df)
    return float(moves["Valor_sinal"].sum()) if not moves.empty else 0.0


def month_summary(df: pd.DataFrame, month: str) -> MonthSummary:
    """Resumo do mês. Saldo = salário − soma das dívidas PAGAS no mês."""
    salary = salary_for_month(df, month)
    debts = debts_for_month(df, month)
    paid = float(payments_for_month(df, month)["Valor"].sum())
    total = float(debts["Valor_mes"].sum()) if not debts.empty else 0.0
    moves = piggy_movements(df)
    piggy_month = float(moves.loc[moves["Mes_Referencia"] == month, "Valor_sinal"].sum()) if not moves.empty else 0.0
    base = salary or 0.0
    return MonthSummary(
        month=month,
        salary=salary,
        total_debts=total,
        paid=paid,
        pending=max(total - paid, 0.0),
        balance=base - paid,
        projected_balance=base - total,
        piggy_bank=piggy_balance(df),
        piggy_month_net=piggy_month,
    )


def monthly_history(df: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por mês com salário, total pago e saldo (para o gráfico de histórico)."""
    months = sorted(
        set(df.loc[df["Registro"].isin([REG_SALARIO, REG_PAGAMENTO]), "Mes_Referencia"]) - {""}
    )
    rows = []
    for m in months:
        salary = salary_for_month(df, m) or 0.0
        paid = float(payments_for_month(df, m)["Valor"].sum())
        rows.append({"Mês": month_label(m), "mes": m, "Salário": salary, "Pago": paid, "Saldo": salary - paid})
    return pd.DataFrame(rows, columns=["Mês", "mes", "Salário", "Pago", "Saldo"])
