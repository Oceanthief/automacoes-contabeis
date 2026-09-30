"""
Monta o ResumoExecucao (período, contagens) a partir dos resultados do
motor de cálculo. Separado do exportador de Excel porque o resumo também
serve pra imprimir no terminal — não faz sentido esse texto depender de
ter gerado um arquivo .xlsx antes.
"""
from __future__ import annotations

from .models import LancamentoContamax, ResumoExecucao, StatusLancamento
from .regras import HISTORICO_APLICACAO, HISTORICO_RESGATE


def montar_resumo(resultados: list[LancamentoContamax]) -> ResumoExecucao:
    if not resultados:
        raise ValueError("Não há resultados para resumir (lista vazia).")

    anos_meses = {(r.data.year, r.data.month) for r in resultados}
    ano, mes = next(iter(anos_meses))
    periodo = f"{ano}-{mes:02d}"

    total_resgates = sum(
        1 for r in resultados if r.status == StatusLancamento.GERADO and r.historico == HISTORICO_RESGATE
    )
    total_aplicacoes = sum(
        1 for r in resultados if r.status == StatusLancamento.GERADO and r.historico == HISTORICO_APLICACAO
    )
    pendencias = sum(
        1
        for r in resultados
        if r.status in (StatusLancamento.SEM_LANCAMENTO_EMPATE, StatusLancamento.ERRO_VALOR_INVALIDO)
    )

    return ResumoExecucao(
        periodo=periodo,
        dias_processados=len(resultados),
        total_aplicacoes=total_aplicacoes,
        total_resgates=total_resgates,
        pendencias=pendencias,
    )


def resumo_como_texto(resumo: ResumoExecucao) -> str:
    return (
        f"Período: {resumo.periodo}\n"
        f"Dias processados: {resumo.dias_processados}\n"
        f"Total de aplicações: {resumo.total_aplicacoes}\n"
        f"Total de resgates: {resumo.total_resgates}\n"
        f"Pendências: {resumo.pendencias}"
    )
