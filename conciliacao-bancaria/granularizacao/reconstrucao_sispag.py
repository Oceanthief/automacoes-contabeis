"""Reconstrução estruturada dos pagamentos por trás de cada SISPAG FORNECEDORES.

Não reimplementa a conciliação: reinvoca as MESMAS funções já auditadas
(leitura.contas_pagas + leitura.extrato_padronizado + conciliacao.motor_sispag)
para obter os objetos Pagamento estruturados que o motor de conciliação já
produz internamente, mas que auditoria_sispag.xlsx só exporta como colunas de
texto concatenado por ';' — inseguras para reconstrução automática (um nome
de fornecedor poderia em tese conter ';', sem garantia de alinhamento
posicional entre as colunas TITULOS/NOMES/VALORES).

Por segurança, o resultado da reinvocação é conferido, SEQUENCIA a SEQUENCIA,
contra o auditoria_sispag.xlsx já publicado (mesmo STATUS, mesmo
TOTAL_ENCONTRADO) antes de ser usado. Qualquer divergência levanta
EstruturaSispagInsuficienteError em vez de prosseguir com dado
potencialmente desatualizado.
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import pandas as pd

from conciliacao.motor_sispag import ResultadoSispag, conciliar
from leitura.contas_pagas import ler_contas_pagas, montar_pagamentos
from leitura.extrato_padronizado import ler_extrato_padronizado, listar_sispag_fornecedores
from normalizacao import valor_para_decimal


class EstruturaSispagInsuficienteError(Exception):
    """ESTRUTURA_SISPAG_INSUFICIENTE: a reconstrução não bate com o que já
    foi auditado — parar em vez de seguir com dado potencialmente
    desatualizado."""


def reconstruir_resultados_sispag(
    caminho_extrato_padronizado: str,
    caminho_contas_pagas: str,
    caminho_auditoria_sispag: str,
) -> list[ResultadoSispag]:
    """Reinvoca a conciliação já validada para obter os Pagamento estruturados
    de cada SISPAG, e confere o resultado contra a auditoria já publicada.
    """
    extrato = ler_extrato_padronizado(caminho_extrato_padronizado)
    sispag_df = listar_sispag_fornecedores(extrato).sort_values("SEQUENCIA")
    sispags = [
        (
            int(l["SEQUENCIA"]),
            l["DATA"].date() if hasattr(l["DATA"], "date") else l["DATA"],
            valor_para_decimal(l["VALOR"]),
        )
        for _, l in sispag_df.iterrows()
    ]

    contas_pagas_df = ler_contas_pagas(caminho_contas_pagas)
    candidatos, _fora_do_escopo = montar_pagamentos(contas_pagas_df)

    resultados = conciliar(sispags, candidatos)

    auditoria_publicada = pd.read_excel(caminho_auditoria_sispag, sheet_name=0)
    _conferir_contra_auditoria_publicada(resultados, auditoria_publicada)

    return resultados


def _conferir_contra_auditoria_publicada(resultados: list[ResultadoSispag], auditoria: pd.DataFrame) -> None:
    auditoria_por_sequencia = {int(r["SEQUENCIA"]): r for _, r in auditoria.iterrows()}

    divergencias = []
    for resultado in resultados:
        linha_publicada = auditoria_por_sequencia.get(resultado.sequencia)
        if linha_publicada is None:
            divergencias.append(f"SEQUENCIA {resultado.sequencia}: não existe na auditoria já publicada.")
            continue

        status_publicado = str(linha_publicada["STATUS"]).strip()
        if status_publicado != resultado.status:
            divergencias.append(
                f"SEQUENCIA {resultado.sequencia}: STATUS recalculado ({resultado.status}) "
                f"diferente do já publicado ({status_publicado})."
            )
            continue

        total_publicado = valor_para_decimal(linha_publicada["TOTAL_ENCONTRADO"])
        if abs(total_publicado - resultado.total_encontrado) > Decimal("0.01"):
            divergencias.append(
                f"SEQUENCIA {resultado.sequencia}: TOTAL_ENCONTRADO recalculado "
                f"({resultado.total_encontrado}) diferente do já publicado ({total_publicado})."
            )

    if divergencias:
        detalhe = "\n".join(divergencias)
        raise EstruturaSispagInsuficienteError(
            "ESTRUTURA_SISPAG_INSUFICIENTE: a reconstrução estruturada dos pagamentos "
            f"diverge da auditoria SISPAG já publicada — parando em vez de prosseguir.\n{detalhe}"
        )
