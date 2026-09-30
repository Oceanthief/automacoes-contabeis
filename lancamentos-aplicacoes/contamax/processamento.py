"""
Orquestração do pipeline: leitura do .xls -> agrupamento por data -> regra
contábil -> lista de LancamentoContamax (um por data).

Este módulo é a única "porta de entrada" pensada para uso externo neste
Prompt 2. Ele não gera Excel nem TXT (fora de escopo) — só devolve os
resultados em memória, prontos para inspeção/teste.
"""
from __future__ import annotations

from decimal import Decimal
from typing import Iterable

from .leitura_xls import ler_extrato
from .models import LancamentoContamax, LinhaComErro, Movimentacao, StatusLancamento
from .regras import agrupar_por_data, aplicar_regra_contamax


def processar_movimentacoes(
    movimentacoes: Iterable[Movimentacao],
    linhas_com_erro: Iterable[LinhaComErro] = (),
) -> list[LancamentoContamax]:
    """
    Aplica agrupamento + regra contábil a uma lista de movimentações já
    normalizadas (independente de terem vindo de um .xls ou de um teste
    sintético — por isso esta função não recebe um caminho de arquivo).

    Qualquer data presente em `linhas_com_erro` sai com
    status=ERRO_VALOR_INVALIDO, mesmo que outras linhas daquele mesmo dia
    tenham sido lidas com sucesso: um total contábil parcialmente calculado
    (por ter descartado sem avisar a linha problemática) seria pior do que
    marcar o dia inteiro como pendente de revisão manual.
    """
    totais = agrupar_por_data(movimentacoes)
    datas_com_erro = {linha.data for linha in linhas_com_erro if linha.data is not None}

    todas_as_datas = set(totais) | datas_com_erro
    resultados: list[LancamentoContamax] = []

    for data in sorted(todas_as_datas):
        total_debito, total_credito = totais.get(data, (Decimal("0.00"), Decimal("0.00")))

        if data in datas_com_erro:
            resultados.append(
                LancamentoContamax(
                    data=data,
                    total_debito=total_debito,
                    total_credito=total_credito,
                    diferenca=total_debito - total_credito,
                    status=StatusLancamento.ERRO_VALOR_INVALIDO,
                )
            )
            continue

        resultados.append(aplicar_regra_contamax(data, total_debito, total_credito))

    return resultados


def processar_extrato_xls(caminho: str) -> list[LancamentoContamax]:
    """Ponto de entrada de conveniência: lê o .xls e já devolve os resultados."""
    movimentacoes, linhas_com_erro = ler_extrato(caminho)
    return processar_movimentacoes(movimentacoes, linhas_com_erro)
