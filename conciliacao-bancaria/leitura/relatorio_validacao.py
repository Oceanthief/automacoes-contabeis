"""Leitura de relatorio_validacao.xlsx (saída da Fase 6)."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

STATUS_GERAL_APROVADO = "APROVADO_PARA_TXT"

COLUNAS_ESPERADAS = [
    "SEQUENCIA_CONTABIL", "SEQUENCIA_BANCO", "DEBITO", "CREDITO", "DATA", "VALOR",
    "NOVA_DESCRICAO", "STATUS_VALIDACAO",
]


class RelatorioValidacaoError(Exception):
    """Erros de leitura/formato/pré-condição do relatório de validação."""


def ler_status_geral(caminho: str | Path) -> str:
    """Lê a aba RESUMO e devolve o valor de STATUS_GERAL."""
    caminho = Path(caminho)
    if not caminho.exists():
        raise RelatorioValidacaoError(f"Arquivo não encontrado: {caminho}")

    resumo = pd.read_excel(caminho, sheet_name="RESUMO")
    linha_status = resumo[resumo.iloc[:, 0] == "STATUS_GERAL"]
    if linha_status.empty:
        raise RelatorioValidacaoError("Aba RESUMO não contém a linha STATUS_GERAL.")
    return str(linha_status.iloc[0, 1]).strip()


def ler_linhas_aprovadas(caminho: str | Path) -> pd.DataFrame:
    """Lê a aba LINHAS_APROVADAS — já filtrada pela Fase 6, mas reconfirmada
    aqui (STATUS_VALIDACAO == APROVADO) por segurança, nunca por confiança cega."""
    caminho = Path(caminho)
    if not caminho.exists():
        raise RelatorioValidacaoError(f"Arquivo não encontrado: {caminho}")

    try:
        df = pd.read_excel(
            caminho, sheet_name="LINHAS_APROVADAS",
            dtype={"DEBITO": str, "CREDITO": str, "TITULO": str},
        )
    except Exception as exc:
        raise RelatorioValidacaoError(f"Não foi possível abrir a aba LINHAS_APROVADAS: {exc}") from exc

    faltantes = [c for c in COLUNAS_ESPERADAS if c not in df.columns]
    if faltantes:
        raise RelatorioValidacaoError(f"Colunas ausentes em LINHAS_APROVADAS: {faltantes}")

    df = df[df["STATUS_VALIDACAO"] == "APROVADO"].reset_index(drop=True)
    return df
