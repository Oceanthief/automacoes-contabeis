"""
Validações executadas antes de gravar o Excel de conferência.

Filosofia: são checagens de sanidade sobre o que o motor de cálculo (já
testado no Prompt 2) produziu. Elas não deveriam falhar em uso normal — mas
existem justamente para pegar um bug futuro (numa alteração de código, ou
num dado de entrada fora do que foi previsto) ANTES que ele vire uma
planilha de conferência com números errados sendo levada pra alguém
decidir em cima. Falhar aqui, de forma barulhenta, é sempre melhor do que
entregar um Excel "quase certo".
"""
from __future__ import annotations

from decimal import Decimal

from .models import LancamentoContamax, StatusLancamento
from .regras import (
    CONTA_APLICACAO_CONTAMAX,
    CONTA_SALDO_CORRENTE,
    HISTORICO_APLICACAO,
    HISTORICO_RESGATE,
)


class ErroCriticoValidacao(Exception):
    """
    Levantada quando alguma checagem de sanidade falha. Contém a lista
    completa de problemas encontrados (não só o primeiro) em `.problemas`,
    para quem for corrigir não precisar rodar o processo várias vezes.
    """

    def __init__(self, problemas: list[str]):
        self.problemas = problemas
        mensagem = "Validação encontrou " + str(len(problemas)) + " problema(s):\n- " + "\n- ".join(problemas)
        super().__init__(mensagem)


def validar_resultados(resultados: list[LancamentoContamax]) -> None:
    """
    Roda todas as checagens pedidas no Prompt 3. Levanta ErroCriticoValidacao
    (interrompendo a geração do Excel) se qualquer uma falhar.
    """
    problemas: list[str] = []

    gerados = [r for r in resultados if r.status == StatusLancamento.GERADO]

    problemas += _checar_datas_duplicadas(gerados)
    problemas += _checar_valores_negativos(resultados)
    problemas += _checar_contas_compativeis_com_historico(gerados)
    problemas += _checar_valor_igual_diferenca_absoluta(gerados)
    problemas += _checar_competencia_unica(resultados)

    if problemas:
        raise ErroCriticoValidacao(problemas)


def _checar_datas_duplicadas(gerados: list[LancamentoContamax]) -> list[str]:
    # Defensivo: o motor de cálculo (contamax.processamento) já agrupa por
    # data e produz no máximo um LancamentoContamax GERADO por data, então
    # isto não deveria disparar em uso normal — mas se um bug futuro quebrar
    # essa garantia, é aqui que ele é pego antes de virar um Excel com duas
    # linhas pra mesma data em LANCAMENTOS.
    datas = [r.data for r in gerados]
    duplicadas = sorted({d for d in datas if datas.count(d) > 1})
    return [
        f"Data duplicada em LANCAMENTOS: {d.strftime('%d/%m/%Y')}" for d in duplicadas
    ]


def _checar_valores_negativos(resultados: list[LancamentoContamax]) -> list[str]:
    problemas = []
    for r in resultados:
        if r.total_debito < 0:
            problemas.append(f"{r.data.strftime('%d/%m/%Y')}: TOTAL_DEBITO negativo ({r.total_debito})")
        if r.total_credito < 0:
            problemas.append(f"{r.data.strftime('%d/%m/%Y')}: TOTAL_CREDITO negativo ({r.total_credito})")
        if r.valor is not None and r.valor < 0:
            problemas.append(f"{r.data.strftime('%d/%m/%Y')}: VALOR negativo ({r.valor})")
    return problemas


def _checar_contas_compativeis_com_historico(gerados: list[LancamentoContamax]) -> list[str]:
    problemas = []
    for r in gerados:
        if r.historico == HISTORICO_RESGATE:
            esperado = (CONTA_APLICACAO_CONTAMAX, CONTA_SALDO_CORRENTE)
        elif r.historico == HISTORICO_APLICACAO:
            esperado = (CONTA_SALDO_CORRENTE, CONTA_APLICACAO_CONTAMAX)
        else:
            problemas.append(
                f"{r.data.strftime('%d/%m/%Y')}: histórico '{r.historico}' não reconhecido pela regra CONTAMAX"
            )
            continue
        if (r.conta_debito, r.conta_credito) != esperado:
            problemas.append(
                f"{r.data.strftime('%d/%m/%Y')}: contas ({r.conta_debito}/{r.conta_credito}) "
                f"incompatíveis com histórico '{r.historico}' (esperado {esperado[0]}/{esperado[1]})"
            )
    return problemas


def _checar_valor_igual_diferenca_absoluta(gerados: list[LancamentoContamax]) -> list[str]:
    problemas = []
    for r in gerados:
        if r.valor != abs(r.diferenca):
            problemas.append(
                f"{r.data.strftime('%d/%m/%Y')}: VALOR ({r.valor}) difere de |DIFERENCA| ({abs(r.diferenca)})"
            )
    return problemas


def _checar_competencia_unica(resultados: list[LancamentoContamax]) -> list[str]:
    # Checagem extra (não pedida explicitamente, mas decorre direto do nome
    # de arquivo pedido: contamax_AAAA-MM.xlsx pressupõe um único mês de
    # competência). Se o extrato processado misturar meses — o que não deveria
    # acontecer num arquivo mensal normal — não dá pra nomear o arquivo com
    # segurança nem pra chamar isso de "o Excel de setembro", por exemplo.
    meses = {(r.data.year, r.data.month) for r in resultados}
    if len(meses) > 1:
        meses_fmt = ", ".join(f"{ano}-{mes:02d}" for ano, mes in sorted(meses))
        return [f"Os resultados abrangem mais de um mês de competência: {meses_fmt}"]
    return []
