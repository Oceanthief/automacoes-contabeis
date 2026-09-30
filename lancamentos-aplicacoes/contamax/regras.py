"""
Núcleo de cálculo: agrupamento por data + aplicação das duas regras contábeis
do CONTAMAX. Não sabe nada sobre arquivos .xls nem .txt — só trabalha com
os tipos definidos em contamax.models. Isso é proposital: mantém a regra de
negócio testável com dados 100% sintéticos, sem precisar montar um Excel
para cada cenário de teste.
"""
from __future__ import annotations

from collections import defaultdict
from datetime import date
from decimal import Decimal
from typing import Iterable

from .models import LancamentoContamax, Movimentacao, StatusLancamento

CONTA_APLICACAO_CONTAMAX = 11023
CONTA_SALDO_CORRENTE = 11145

HISTORICO_RESGATE = "VLR REF. RESGATE CONTAMAX"
HISTORICO_APLICACAO = "VLR REF. APLICAÇÃO CONTAMAX"


def agrupar_por_data(
    movimentacoes: Iterable[Movimentacao],
) -> dict[date, tuple[Decimal, Decimal]]:
    """
    Soma débito e crédito por data.

    Retorna um dict {data: (total_debito, total_credito)}. A soma usa
    Decimal em todo o caminho — cada Movimentacao já chega com valores em
    Decimal (contamax.valores cuida disso na leitura), então não há
    conversão de tipo aqui, só acumulação.
    """
    totais: dict[date, list[Decimal]] = defaultdict(lambda: [Decimal("0.00"), Decimal("0.00")])
    for mov in movimentacoes:
        par = totais[mov.data]
        par[0] += mov.debito
        par[1] += mov.credito
    return {data: (debito, credito) for data, (debito, credito) in totais.items()}


def aplicar_regra_contamax(
    data: date, total_debito: Decimal, total_credito: Decimal
) -> LancamentoContamax:
    """
    Aplica a regra contábil confirmada na especificação (seção 3) a um único
    dia já totalizado.

    Caso de empate (decisão 5.1 da especificação): não gera lançamento.
    O dia ainda volta como um LancamentoContamax, só que com
    status=SEM_LANCAMENTO_EMPATE e os campos de lançamento em None — isso é
    a forma "controlada" de tratar o caso pedido no Prompt 2: nada é
    inventado, e o caso fica claramente identificável no resultado (em vez
    de a linha simplesmente desaparecer sem explicação).
    """
    diferenca = total_debito - total_credito

    if diferenca > 0:
        return LancamentoContamax(
            data=data,
            total_debito=total_debito,
            total_credito=total_credito,
            diferenca=diferenca,
            status=StatusLancamento.GERADO,
            conta_debito=CONTA_APLICACAO_CONTAMAX,
            conta_credito=CONTA_SALDO_CORRENTE,
            historico=HISTORICO_RESGATE,
            valor=diferenca,
        )

    if diferenca < 0:
        valor = -diferenca
        return LancamentoContamax(
            data=data,
            total_debito=total_debito,
            total_credito=total_credito,
            diferenca=diferenca,
            status=StatusLancamento.GERADO,
            conta_debito=CONTA_SALDO_CORRENTE,
            conta_credito=CONTA_APLICACAO_CONTAMAX,
            historico=HISTORICO_APLICACAO,
            valor=valor,
        )

    # diferenca == 0: total_debito == total_credito (decisão 5.1)
    return LancamentoContamax(
        data=data,
        total_debito=total_debito,
        total_credito=total_credito,
        diferenca=diferenca,
        status=StatusLancamento.SEM_LANCAMENTO_EMPATE,
    )
