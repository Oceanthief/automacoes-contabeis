"""Leitura e localização de cabeçalho do Extrato Itaú.

A Fase 0 confirmou que o cabeçalho não está em posição fixa (as primeiras
linhas do extrato são metadados: Atualização, Nome, Agência, Conta, Período).
Este módulo localiza o cabeçalho pelo conteúdo, nunca por número de linha.
"""
from pathlib import Path

import pandas as pd

COLUNAS_ESPERADAS = [
    "Data",
    "Lançamento",
    "Razão Social",
    "CPF/CNPJ",
    "Valor (R$)",
    "Saldo (R$)",
]

COLUNA_VALOR = "Valor (R$)"


class ExtratoItauError(Exception):
    """Erros de leitura/formato específicos do Extrato Itaú."""


def _localizar_linha_cabecalho(bruto: pd.DataFrame) -> int:
    """Varre as linhas até achar a que contém todas as colunas esperadas.

    Comparação célula a célula (não string inteira da linha) para não
    depender de quantas colunas existem à direita ou da ordem exata.
    """
    for indice, linha in bruto.iterrows():
        valores = {str(v).strip() for v in linha.tolist()}
        if set(COLUNAS_ESPERADAS).issubset(valores):
            return indice
    raise ExtratoItauError(
        "Não foi possível localizar a linha de cabeçalho do extrato. "
        f"Colunas esperadas: {COLUNAS_ESPERADAS}"
    )


def ler_extrato(caminho: str | Path) -> pd.DataFrame:
    """Lê o Extrato Itaú e devolve um DataFrame com o cabeçalho correto.

    Abre em modo leitura apenas — o arquivo original nunca é reescrito.
    """
    caminho = Path(caminho)
    if not caminho.exists():
        raise ExtratoItauError(f"Arquivo não encontrado: {caminho}")

    try:
        planilha = pd.ExcelFile(caminho)
    except Exception as exc:  # arquivo corrompido, formato inválido, etc.
        raise ExtratoItauError(f"Não foi possível abrir o arquivo '{caminho}': {exc}") from exc

    if not planilha.sheet_names:
        raise ExtratoItauError(f"Arquivo sem nenhuma aba: {caminho}")

    bruto = planilha.parse(planilha.sheet_names[0], header=None)
    linha_cabecalho = _localizar_linha_cabecalho(bruto)
    df = planilha.parse(planilha.sheet_names[0], header=linha_cabecalho)

    faltantes = [c for c in COLUNAS_ESPERADAS if c not in df.columns]
    if faltantes:
        raise ExtratoItauError(f"Colunas ausentes no extrato: {faltantes}")

    return df


def filtrar_movimentacoes(df: pd.DataFrame) -> pd.DataFrame:
    """Mantém só linhas com Valor (R$) preenchido.

    Isso elimina as linhas de saldo (SALDO ANTERIOR, SALDO TOTAL DISPONÍVEL
    DIA) sem depender do texto do lançamento — a Fase 0 confirmou que essas
    linhas são exatamente as que têm Valor (R$) nulo, 100% dos casos na
    amostra de agosto/2026.
    """
    filtrado = df[df[COLUNA_VALOR].notna()].reset_index(drop=True)
    return filtrado
