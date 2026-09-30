"""Leitura do Excel de auditoria gerado por conciliar_sispag.py."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

COLUNAS_ESPERADAS = [
    "SEQUENCIA",
    "DATA",
    "VALOR_SISPAG",
    "STATUS",
    "TOTAL_ENCONTRADO",
    "DIFERENCA",
    "QTD_PAGAMENTOS",
    "TITULOS",
    "NOMES",
    "NF_NUM",
    "VALORES",
    "FORMA_PGTO",
    "HISTORICOS",
]


class AuditoriaSispagError(Exception):
    """Erros de leitura/formato do Excel de auditoria SISPAG."""


def ler_auditoria_sispag(caminho: str | Path) -> pd.DataFrame:
    caminho = Path(caminho)
    if not caminho.exists():
        raise AuditoriaSispagError(f"Arquivo não encontrado: {caminho}")

    try:
        df = pd.read_excel(caminho, sheet_name=0)
    except Exception as exc:
        raise AuditoriaSispagError(f"Não foi possível abrir '{caminho}': {exc}") from exc

    faltantes = [c for c in COLUNAS_ESPERADAS if c not in df.columns]
    if faltantes:
        raise AuditoriaSispagError(f"Colunas ausentes na auditoria: {faltantes}")

    return df
