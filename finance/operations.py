"""Regras de negócio que ALTERAM os dados de um usuário.

Todas as funções são puras: recebem um DataFrame, devolvem um NOVO DataFrame e não
tocam em Streamlit nem no Excel. Isso facilita testes e permite desfazer a alteração
(rollback) se o salvamento na planilha falhar.
"""

from __future__ import annotations

import uuid
from datetime import datetime

import pandas as pd

from finance.repository import (
    COLUMNS,
    REG_COFRINHO,
    REG_DIVIDA,
    REG_PAGAMENTO,
    REG_SALARIO,
    normalize,
)

TIPO_RECORRENTE = "Recorrente"
TIPO_PARCELADA = "Parcelada"
TIPO_DEPOSITO = "Depósito"
TIPO_RETIRADA = "Retirada"
STATUS_PAGO = "Pago"
STATUS_PENDENTE = "Pendente"
STATUS_QUITADA = "Quitada"
STATUS_ENCERRADA = "Encerrada"
SIM, NAO = "SIM", "NAO"


class ValidationError(ValueError):
    """Entrada inválida do usuário (mensagem pronta para exibir)."""


def _new_id() -> str:
    return uuid.uuid4().hex[:8]


def _append(df: pd.DataFrame, **values) -> pd.DataFrame:
    row = {col: None for col in COLUMNS}
    row.update(values)
    new = pd.DataFrame([row], columns=COLUMNS)
    if df.empty:
        return normalize(new)
    return normalize(pd.concat([df, new], ignore_index=True))


def _debt_index(df: pd.DataFrame, debt_id: str) -> int:
    mask = (df["Registro"] == REG_DIVIDA) & (df["ID"] == debt_id)
    if not mask.any():
        raise ValidationError("Dívida não encontrada. Recarregue a planilha.")
    return int(df.index[mask][0])


def _payment_mask(df: pd.DataFrame, debt_id: str, month: str) -> pd.Series:
    return (
        (df["Registro"] == REG_PAGAMENTO)
        & (df["ID_Divida"] == debt_id)
        & (df["Mes_Referencia"] == month)
    )


# --------------------------------------------------------------------------- #
# Salário
# --------------------------------------------------------------------------- #
def set_salary(df: pd.DataFrame, month: str, value: float, now: datetime | None = None) -> pd.DataFrame:
    """Cria ou atualiza o salário do mês (uma linha SALARIO por mês)."""
    if value is None or value < 0:
        raise ValidationError("O salário deve ser um valor maior ou igual a zero.")
    now = now or datetime.now()
    df = df.copy()
    mask = (df["Registro"] == REG_SALARIO) & (df["Mes_Referencia"] == month)
    if mask.any():
        idx = df.index[mask]
        df.loc[idx[0], ["Valor", "Data"]] = [float(value), pd.Timestamp(now)]
        return normalize(df.drop(index=idx[1:]))  # remove duplicatas eventuais do mesmo mês
    return _append(
        df, Registro=REG_SALARIO, ID=_new_id(), Mes_Referencia=month, Data=now,
        Descricao="Salário", Valor=float(value),
    )


# --------------------------------------------------------------------------- #
# Dívidas
# --------------------------------------------------------------------------- #
def add_debt(
    df: pd.DataFrame,
    name: str,
    kind: str,
    installment: float,
    start_month: str,
    category: str = "",
    total_installments: int | None = None,
    paid_installments: int = 0,
    now: datetime | None = None,
) -> pd.DataFrame:
    """Cadastra uma dívida recorrente (sem fim) ou parcelada (com total de parcelas)."""
    name = (name or "").strip()
    if not name:
        raise ValidationError("Informe o nome da dívida.")
    if kind not in (TIPO_RECORRENTE, TIPO_PARCELADA):
        raise ValidationError("Tipo de dívida inválido.")
    if installment is None or installment <= 0:
        raise ValidationError("O valor da parcela deve ser maior que zero.")
    paid_installments = int(paid_installments or 0)
    if paid_installments < 0:
        raise ValidationError("Parcelas já pagas não pode ser negativo.")

    if kind == TIPO_PARCELADA:
        if not total_installments or int(total_installments) < 1:
            raise ValidationError("Informe a quantidade total de parcelas (mínimo 1).")
        total_installments = int(total_installments)
        if paid_installments >= total_installments:
            raise ValidationError("Parcelas já pagas deve ser menor que o total de parcelas.")
    else:
        total_installments = None  # recorrente = sem fim definido (célula vazia no Excel)

    return _append(
        df, Registro=REG_DIVIDA, ID=_new_id(), Mes_Referencia=start_month,
        Data=now or datetime.now(), Descricao=name, Categoria=(category or "").strip(),
        Tipo=kind, Valor=float(installment), Parcelas_Total=total_installments,
        Parcelas_Pagas=paid_installments, Status_Mes=STATUS_PENDENTE, Ativa=SIM,
    )


def pay_debt(df: pd.DataFrame, debt_id: str, month: str, now: datetime | None = None) -> pd.DataFrame:
    """Marca a parcela do mês como paga.

    * incrementa ``Parcelas_Pagas``;
    * registra uma linha PAGAMENTO (é ela que debita o saldo do mês);
    * encerra automaticamente a dívida parcelada quando a última parcela é paga.
    """
    df = df.copy()
    idx = _debt_index(df, debt_id)
    debt = df.loc[idx]
    if debt["Ativa"] != SIM:
        raise ValidationError(f"A dívida '{debt['Descricao']}' já está quitada/encerrada.")
    if _payment_mask(df, debt_id, month).any():
        raise ValidationError(f"A parcela de '{debt['Descricao']}' já foi paga neste mês.")

    paid = int(0 if pd.isna(debt["Parcelas_Pagas"]) else debt["Parcelas_Pagas"]) + 1
    df.loc[idx, "Parcelas_Pagas"] = float(paid)
    total = debt["Parcelas_Total"]
    if debt["Tipo"] == TIPO_PARCELADA and not pd.isna(total) and paid >= int(total):
        df.loc[idx, "Ativa"] = NAO

    parcel_label = f"{paid}/{int(total)}" if not pd.isna(total) else f"{paid}"
    return _append(
        df, Registro=REG_PAGAMENTO, ID=_new_id(), Mes_Referencia=month, Data=now or datetime.now(),
        Descricao=f"{debt['Descricao']} (parcela {parcel_label})", Categoria=debt["Categoria"],
        Tipo=debt["Tipo"], Valor=float(debt["Valor"]), ID_Divida=debt_id,
    )


def undo_payment(df: pd.DataFrame, debt_id: str, month: str) -> pd.DataFrame:
    """Desfaz o pagamento do mês (para corrigir cliques errados)."""
    df = df.copy()
    idx = _debt_index(df, debt_id)
    mask = _payment_mask(df, debt_id, month)
    if not mask.any():
        raise ValidationError("Não há pagamento desta dívida no mês selecionado.")
    paid = max(int(0 if pd.isna(df.loc[idx, "Parcelas_Pagas"]) else df.loc[idx, "Parcelas_Pagas"]) - 1, 0)
    df.loc[idx, "Parcelas_Pagas"] = float(paid)
    total = df.loc[idx, "Parcelas_Total"]
    if df.loc[idx, "Tipo"] == TIPO_PARCELADA and not pd.isna(total) and paid < int(total):
        df.loc[idx, "Ativa"] = SIM  # volta a ficar ativa se tinha sido quitada por esse pagamento
    first_payment = df.index[mask][0]
    return normalize(df.drop(index=first_payment))


def close_debt(df: pd.DataFrame, debt_id: str) -> pd.DataFrame:
    """Encerra uma dívida (ex.: cancelou a assinatura). O histórico de pagamentos é mantido."""
    df = df.copy()
    idx = _debt_index(df, debt_id)
    df.loc[idx, "Ativa"] = NAO
    df.loc[idx, "Status_Mes"] = STATUS_ENCERRADA
    return df


# --------------------------------------------------------------------------- #
# Cofrinho
# --------------------------------------------------------------------------- #
def move_piggy_bank(
    df: pd.DataFrame,
    kind: str,
    value: float,
    month: str,
    description: str = "",
    current_balance: float = 0.0,
    now: datetime | None = None,
) -> pd.DataFrame:
    """Registra depósito ou retirada no cofrinho (valor sempre positivo + tipo)."""
    if kind not in (TIPO_DEPOSITO, TIPO_RETIRADA):
        raise ValidationError("Tipo de movimentação inválido.")
    if value is None or value <= 0:
        raise ValidationError("O valor deve ser maior que zero.")
    if kind == TIPO_RETIRADA and value > current_balance + 1e-9:
        raise ValidationError("Saldo insuficiente no cofrinho para essa retirada.")
    return _append(
        df, Registro=REG_COFRINHO, ID=_new_id(), Mes_Referencia=month, Data=now or datetime.now(),
        Descricao=(description or kind).strip(), Tipo=kind, Valor=float(value),
    )


# --------------------------------------------------------------------------- #
# Status do mês (coluna informativa gravada no Excel)
# --------------------------------------------------------------------------- #
def refresh_month_status(df: pd.DataFrame, month: str) -> pd.DataFrame:
    """Atualiza a coluna ``Status_Mes`` das dívidas com base nos pagamentos de ``month``."""
    df = df.copy()
    paid_ids = set(
        df.loc[(df["Registro"] == REG_PAGAMENTO) & (df["Mes_Referencia"] == month), "ID_Divida"]
    )
    for idx in df.index[df["Registro"] == REG_DIVIDA]:
        row = df.loc[idx]
        if row["ID"] in paid_ids:
            status = STATUS_PAGO
        elif row["Ativa"] == SIM:
            status = STATUS_PENDENTE
        elif row["Tipo"] == TIPO_PARCELADA and row["Parcelas_Pagas"] >= (row["Parcelas_Total"] or 0):
            status = STATUS_QUITADA
        else:
            status = STATUS_ENCERRADA
        df.loc[idx, "Status_Mes"] = status
    return df
