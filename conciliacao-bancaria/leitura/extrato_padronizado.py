"""Leitura do Extrato Itaú já padronizado pela Fase 1 (8 colunas fixas).

Diferente de leitura/extrato_itau.py (que lê o extrato bruto do banco), este
módulo lê a saída de fase1_padronizar_extrato.py — já limpa, sem cabeçalho a
localizar dinamicamente.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

# As duas colunas "SEM INFORMACAO" da Fase 1 viram "SEM INFORMACAO" e
# "SEM INFORMACAO.1" quando o pandas lê o Excel (ele desambigua nomes
# repetidos sozinho) — por isso não entram nesta checagem por nome.
COLUNAS_ESPERADAS = ["DEBITO", "CREDITO", "DATA", "VALOR", "DESCRICAO", "SEQUENCIA"]

DESCRICAO_SISPAG = "SISPAG FORNECEDORES"


class ExtratoPadronizadoError(Exception):
    """Erros de leitura/formato do Extrato já padronizado."""


def ler_extrato_padronizado(caminho: str | Path) -> pd.DataFrame:
    caminho = Path(caminho)
    if not caminho.exists():
        raise ExtratoPadronizadoError(f"Arquivo não encontrado: {caminho}")

    try:
        # DEBITO/CREDITO são códigos de conta, não números — sem dtype=str
        # explícito, o pandas infere a coluna como float64 sozinho e descarta
        # o zero à esquerda de contas como "00248" (bug real já detectado
        # numa fase anterior: leia a ressalva no changelog do projeto).
        df = pd.read_excel(caminho, sheet_name=0, dtype={"DEBITO": str, "CREDITO": str})
    except Exception as exc:
        raise ExtratoPadronizadoError(f"Não foi possível abrir '{caminho}': {exc}") from exc

    faltantes = [c for c in COLUNAS_ESPERADAS if c not in df.columns]
    if faltantes:
        raise ExtratoPadronizadoError(f"Colunas ausentes no extrato padronizado: {faltantes}")

    return df


def listar_sispag_fornecedores(df: pd.DataFrame) -> pd.DataFrame:
    """Filtra as linhas cuja DESCRICAO é exatamente 'SISPAG FORNECEDORES'."""
    mascara = df["DESCRICAO"].astype(str).str.strip() == DESCRICAO_SISPAG
    return df.loc[mascara].reset_index(drop=True)
