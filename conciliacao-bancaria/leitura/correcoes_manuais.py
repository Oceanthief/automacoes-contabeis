"""Leitura do arquivo de correções manuais (SEM_DE_PARA/AMBIGUO revisados por
um humano depois da Fase 5).

Decisão combinada com o usuário (Opção B): nunca editar saida_contabil.xlsx
diretamente — as correções ficam numa camada separada e rastreável, aplicada
por enriquecimento/validacao antes da validação final.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

COLUNAS_ESPERADAS = ["SEQUENCIA_CONTABIL", "DEBITO_MANUAL", "CREDITO_MANUAL", "OBSERVACAO"]

# Todo código de conta usado neste projeto até agora tem 5 dígitos (com zero
# à esquerda quando necessário). Códigos aqui são a exceção confirmada
# explicitamente pelo usuário — nunca adicionados por inferência.
CODIGOS_SEM_PADDING = {"8685"}


class CorrecoesManuaisError(Exception):
    """Erros de leitura/formato do arquivo de correções manuais."""


def ler_correcoes_manuais(caminho: str | Path) -> pd.DataFrame:
    caminho = Path(caminho)
    if not caminho.exists():
        raise CorrecoesManuaisError(f"Arquivo não encontrado: {caminho}")

    try:
        # DEBITO_MANUAL/CREDITO_MANUAL são códigos de conta, não números —
        # sem dtype=str aqui, o pandas infere float64 e derruba o zero à
        # esquerda (bug já visto duas vezes neste projeto).
        df = pd.read_excel(
            caminho, sheet_name=0, dtype={"DEBITO_MANUAL": str, "CREDITO_MANUAL": str}
        )
    except Exception as exc:
        raise CorrecoesManuaisError(f"Não foi possível abrir '{caminho}': {exc}") from exc

    faltantes = [c for c in COLUNAS_ESPERADAS if c not in df.columns]
    if faltantes:
        raise CorrecoesManuaisError(f"Colunas ausentes em correcoes_manuais: {faltantes}")

    df = df[df["SEQUENCIA_CONTABIL"].notna()].reset_index(drop=True)
    if len(df) == 0:
        raise CorrecoesManuaisError("Nenhuma correção encontrada no arquivo.")

    return df


def normalizar_codigo_conta_manual(valor) -> str:
    """Aplica zero à esquerda até 5 dígitos, exceto para CODIGOS_SEM_PADDING."""
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto == "" or texto.lower() == "nan":
        return ""
    if texto in CODIGOS_SEM_PADDING:
        return texto
    if texto.isdigit() and len(texto) < 5:
        return texto.zfill(5)
    return texto
