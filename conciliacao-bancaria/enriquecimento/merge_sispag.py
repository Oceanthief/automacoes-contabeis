"""Preenche NOVA DESCRICAO dos SISPAG FORNECEDORES a partir de auditoria_sispag.xlsx.

Merge controlado por SEQUENCIA entre saida_enriquecida.xlsx (Fase de
enriquecimento) e auditoria_sispag.xlsx (motor de conciliação). Não refaz a
conciliação — só cruza dois resultados já validados, com checagem cruzada de
DATA e VALOR antes de aceitar qualquer vínculo.
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Optional

import pandas as pd

from normalizacao import normalizar_texto_classificacao, valor_para_decimal

DESCRICAO_SISPAG_FORNECEDORES = "SISPAG FORNECEDORES"
STATUS_PREENCHIVEIS = {"EXATO_1X1", "EXATO_SOMA"}
TOLERANCIA_VALOR = Decimal("0.01")


class MergeSispagError(Exception):
    """Erros de integridade do merge (ex.: SEQUENCIA duplicada na auditoria)."""


@dataclass
class ResultadoMerge:
    """Resultado do cruzamento de uma linha SISPAG FORNECEDORES."""

    sequencia: int
    preenchido: bool
    nova_descricao: Optional[str]
    motivo: Optional[str]  # None quando preenchido=True


def _dedupe_nomes(nomes_str: str) -> str:
    """Separa por ';', remove espaços excedentes, remove duplicados por texto
    normalizado (mantendo a PRIMEIRA grafia original encontrada), preserva a
    ordem original e junta de novo com '; '.
    """
    partes = [p.strip() for p in str(nomes_str).split(";")]
    partes = [p for p in partes if p]

    vistos: set[str] = set()
    resultado: list[str] = []
    for nome in partes:
        chave = normalizar_texto_classificacao(nome)
        if chave not in vistos:
            vistos.add(chave)
            resultado.append(nome)

    return "; ".join(resultado)


def _to_date(valor):
    return valor.date() if hasattr(valor, "date") else valor


def processar(saida_df: pd.DataFrame, auditoria_df: pd.DataFrame) -> list[ResultadoMerge]:
    """Para cada linha SISPAG FORNECEDORES de saida_df, tenta achar e validar
    a linha correspondente em auditoria_df pela SEQUENCIA (chave principal).
    """
    sequencias_auditoria = list(auditoria_df["SEQUENCIA"].astype(int))
    if len(sequencias_auditoria) != len(set(sequencias_auditoria)):
        raise MergeSispagError("auditoria_sispag.xlsx tem SEQUENCIA duplicada — dado inválido.")

    auditoria_por_sequencia = {
        int(linha["SEQUENCIA"]): linha for _, linha in auditoria_df.iterrows()
    }

    resultados: list[ResultadoMerge] = []

    for _, linha in saida_df.iterrows():
        if str(linha["DESCRICAO"]).strip() != DESCRICAO_SISPAG_FORNECEDORES:
            continue

        sequencia = int(linha["SEQUENCIA"])
        auditoria_linha = auditoria_por_sequencia.get(sequencia)

        if auditoria_linha is None:
            resultados.append(
                ResultadoMerge(sequencia, False, None, "SEQUENCIA sem correspondência na auditoria")
            )
            continue

        data_saida = _to_date(linha["DATA"])
        data_auditoria = _to_date(auditoria_linha["DATA"])
        if data_saida != data_auditoria:
            resultados.append(
                ResultadoMerge(
                    sequencia, False, None,
                    f"Divergência de DATA (saida={data_saida}, auditoria={data_auditoria})",
                )
            )
            continue

        valor_saida = valor_para_decimal(linha["VALOR"])
        valor_auditoria = valor_para_decimal(auditoria_linha["VALOR_SISPAG"])
        if abs(valor_saida - valor_auditoria) > TOLERANCIA_VALOR:
            resultados.append(
                ResultadoMerge(
                    sequencia, False, None,
                    f"Divergência de VALOR (saida={valor_saida}, auditoria={valor_auditoria})",
                )
            )
            continue

        status = str(auditoria_linha["STATUS"]).strip()
        if status not in STATUS_PREENCHIVEIS:
            resultados.append(
                ResultadoMerge(sequencia, False, None, f"STATUS não preenchível automaticamente ({status})")
            )
            continue

        nomes_raw = auditoria_linha.get("NOMES", "")
        nomes_raw = "" if pd.isna(nomes_raw) else str(nomes_raw)
        nova_descricao = _dedupe_nomes(nomes_raw)

        if not nova_descricao:
            resultados.append(
                ResultadoMerge(sequencia, False, None, "NOMES vazio na auditoria apesar do STATUS preenchível")
            )
            continue

        resultados.append(ResultadoMerge(sequencia, True, nova_descricao, None))

    return resultados
