"""Leitura do Excel de DE-PARA contábil (aba DE_PARA)."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

COLUNAS_ESPERADAS = [
    "CHAVE_MATCH",
    "TIPO_MOVIMENTO",
    "TIPO_MATCH",
    "CONTA_CONTABIL",
    "PRIORIDADE",
    "ATIVO",
]
# HISTORICO_PADRAO e OBSERVACAO também são esperadas, mas são opcionais
# linha a linha (algumas regras legítimas não têm as duas preenchidas).

ABA_REGRAS = "DE_PARA"


class DeParaError(Exception):
    """Erros de leitura/formato do arquivo de DE-PARA."""


def ler_de_para(caminho: str | Path) -> pd.DataFrame:
    """Lê a aba DE_PARA e remove linhas totalmente em branco (separadores
    visuais do Excel) — nunca modifica o arquivo original."""
    caminho = Path(caminho)
    if not caminho.exists():
        raise DeParaError(f"Arquivo não encontrado: {caminho}")

    try:
        df = pd.read_excel(caminho, sheet_name=ABA_REGRAS, dtype={"CONTA_CONTABIL": str})
    except Exception as exc:
        raise DeParaError(f"Não foi possível abrir a aba '{ABA_REGRAS}' de '{caminho}': {exc}") from exc

    faltantes = [c for c in COLUNAS_ESPERADAS if c not in df.columns]
    if faltantes:
        raise DeParaError(f"Colunas ausentes no DE-PARA: {faltantes}")

    # uma linha sem CHAVE_MATCH não é regra nenhuma (separador visual)
    df = df[df["CHAVE_MATCH"].notna()].reset_index(drop=True)

    if len(df) == 0:
        raise DeParaError("Nenhuma regra encontrada na aba DE_PARA (arquivo vazio?).")

    return df
