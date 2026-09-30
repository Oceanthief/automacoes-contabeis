"""Motor de conciliação de lançamentos SISPAG FORNECEDORES.

Explica cada SISPAG do extrato padronizado (leitura/extrato_padronizado.py)
usando os pagamentos únicos reconstruídos de Contas Pagas
(leitura/contas_pagas.py). Não reescreve nenhuma das duas camadas — só as
consome.

ALGORITMO
---------
Etapa 1 (match individual): por data, procura pagamentos ainda disponíveis
  com o mesmo valor exato do SISPAG.
    - 0 candidatos  -> segue para a Etapa 2.
    - 1 candidato   -> EXATO_1X1 (consome o pagamento do pool).
    - 2+ candidatos -> AMBIGUO (não consome nada — não dá pra saber qual é).

Etapa 2 (combinação / subset-sum): quando não há match individual, procura
  subconjuntos de 2+ pagamentos ainda disponíveis daquela data cuja soma
  bate exatamente com o valor do SISPAG. Implementado como backtracking com
  poda:
    - candidatos maiores que o alvo são descartados de cara;
    - candidatos ordenados por valor decrescente;
    - soma de sufixo pré-calculada corta ramos que nunca vão alcançar o alvo;
    - a busca para assim que encontra 2 soluções distintas — o suficiente
      para provar ambiguidade, sem precisar enumerar o espaço todo;
    - uma trava de nós visitados evita explosão combinatória em dias com
      muitos pagamentos (ver LIMITE_NOS_BACKTRACKING).
    - 0 soluções    -> NAO_ENCONTRADO.
    - 1 solução     -> EXATO_SOMA (consome os pagamentos usados).
    - 2+ soluções   -> AMBIGUO (não consome nada).

DECISÃO DE PROJETO — ordem de processamento
--------------------------------------------
Os SISPAG de uma mesma data são processados na ordem recebida (recomenda-se
SEQUENCIA do extrato, ou seja, ordem cronológica). Um pagamento consumido
por um SISPAG anterior no mesmo dia deixa de estar disponível para os
seguintes. Isso é intencional (nenhum pagamento pode alimentar dois SISPAG),
mas é uma dependência de ordem: em tese, processar os SISPAG de um mesmo dia
em outra ordem poderia produzir uma alocação diferente. Como a suíte de
match individual sempre prioriza a correspondência mais "óbvia" primeiro,
o risco prático é baixo, mas fica registrado.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from leitura.contas_pagas import Pagamento
from normalizacao import para_centavos

STATUS_EXATO_1X1 = "EXATO_1X1"
STATUS_EXATO_SOMA = "EXATO_SOMA"
STATUS_AMBIGUO = "AMBIGUO"
STATUS_NAO_ENCONTRADO = "NAO_ENCONTRADO"

MAX_SOLUCOES_PARA_PROVAR_AMBIGUIDADE = 2
LIMITE_NOS_BACKTRACKING = 200_000  # trava de segurança contra explosão combinatória


@dataclass
class ResultadoSispag:
    """Resultado da conciliação de um lançamento SISPAG."""

    sequencia: int
    data: date
    valor_sispag: Decimal
    status: str
    pagamentos: list[Pagamento] = field(default_factory=list)
    # só preenchido quando status == AMBIGUO: cada item é uma combinação
    # candidata (1 pagamento, no caso de ambiguidade individual; 2+, no
    # caso de ambiguidade de soma).
    alternativas: list[list[Pagamento]] = field(default_factory=list)

    @property
    def total_encontrado(self) -> Decimal:
        if self.status in (STATUS_EXATO_1X1, STATUS_EXATO_SOMA):
            return sum((p.valor_classificado for p in self.pagamentos), start=Decimal("0.00"))
        return Decimal("0.00")

    @property
    def diferenca(self) -> Decimal:
        return self.valor_sispag - self.total_encontrado


class _BuscaInterrompidaPorLimite(Exception):
    """Sinaliza que o backtracking bateu no limite de segurança de nós visitados."""


def _buscar_combinacoes(
    candidatos: list[tuple[Pagamento, int]],
    alvo_centavos: int,
    max_solucoes: int = MAX_SOLUCOES_PARA_PROVAR_AMBIGUIDADE,
) -> list[list[Pagamento]]:
    """Backtracking com poda para subset-sum exato (valores em centavos).

    `candidatos` já deve vir filtrado para valor <= alvo (poda de quem chama).
    Retorna até `max_solucoes` combinações de 2+ pagamentos cuja soma bate
    exatamente com `alvo_centavos`, parando assim que acha o suficiente para
    provar ambiguidade.
    """
    candidatos_ordenados = sorted(candidatos, key=lambda par: par[1], reverse=True)
    n = len(candidatos_ordenados)

    soma_sufixo = [0] * (n + 1)
    for i in range(n - 1, -1, -1):
        soma_sufixo[i] = soma_sufixo[i + 1] + candidatos_ordenados[i][1]

    solucoes: list[list[Pagamento]] = []
    nos_visitados = 0

    def backtrack(inicio: int, restante: int, combo_atual: list[Pagamento]) -> None:
        nonlocal nos_visitados
        if len(solucoes) >= max_solucoes:
            return

        nos_visitados += 1
        if nos_visitados > LIMITE_NOS_BACKTRACKING:
            raise _BuscaInterrompidaPorLimite()

        if restante == 0:
            if len(combo_atual) >= 2:
                solucoes.append(list(combo_atual))
            return
        if inicio >= n or soma_sufixo[inicio] < restante:
            return  # poda: nem tudo que resta (mesmo somado) alcança o alvo

        for i in range(inicio, n):
            valor = candidatos_ordenados[i][1]
            if valor > restante:
                continue  # poda: item maior que o que falta
            if soma_sufixo[i] < restante:
                break  # sufixo decrescente monótono: daqui pra frente nunca alcança
            combo_atual.append(candidatos_ordenados[i][0])
            backtrack(i + 1, restante - valor, combo_atual)
            combo_atual.pop()
            if len(solucoes) >= max_solucoes:
                return

    try:
        backtrack(0, alvo_centavos, [])
    except _BuscaInterrompidaPorLimite:
        # Conservador: melhor reportar NAO_ENCONTRADO (nada foi decidido)
        # do que devolver um resultado parcial que pareça definitivo.
        return []

    return solucoes


def _conciliar_um_sispag(
    sequencia: int,
    data_sispag: date,
    valor_sispag: Decimal,
    pool_disponivel: list[Pagamento],
) -> ResultadoSispag:
    valor_centavos = para_centavos(valor_sispag)

    # --- Etapa 1: match individual ---
    candidatos_individuais = [
        p for p in pool_disponivel if para_centavos(p.valor_classificado) == valor_centavos
    ]

    if len(candidatos_individuais) == 1:
        pagamento = candidatos_individuais[0]
        pool_disponivel.remove(pagamento)
        return ResultadoSispag(sequencia, data_sispag, valor_sispag, STATUS_EXATO_1X1, [pagamento])

    if len(candidatos_individuais) >= 2:
        return ResultadoSispag(
            sequencia,
            data_sispag,
            valor_sispag,
            STATUS_AMBIGUO,
            alternativas=[[p] for p in candidatos_individuais],
        )

    # --- Etapa 2: combinação (subset-sum) ---
    candidatos_para_soma = [
        (p, para_centavos(p.valor_classificado))
        for p in pool_disponivel
        if para_centavos(p.valor_classificado) <= valor_centavos
    ]

    solucoes = _buscar_combinacoes(candidatos_para_soma, valor_centavos)

    if len(solucoes) == 1:
        combinacao = solucoes[0]
        for pagamento in combinacao:
            pool_disponivel.remove(pagamento)
        return ResultadoSispag(sequencia, data_sispag, valor_sispag, STATUS_EXATO_SOMA, combinacao)

    if len(solucoes) >= 2:
        return ResultadoSispag(
            sequencia, data_sispag, valor_sispag, STATUS_AMBIGUO, alternativas=solucoes
        )

    return ResultadoSispag(sequencia, data_sispag, valor_sispag, STATUS_NAO_ENCONTRADO)


def conciliar(
    sispags: list[tuple[int, date, Decimal]], pagamentos: list[Pagamento]
) -> list[ResultadoSispag]:
    """Concilia uma lista de SISPAG contra os pagamentos disponíveis.

    sispags: lista de (sequencia, data, valor_sispag) — processados na ordem
      recebida; recomenda-se ordenar por SEQUENCIA antes de chamar.
    pagamentos: pagamentos únicos candidatos (ex.: já filtrados para a conta
      CONTA_ALVO por leitura.contas_pagas.montar_pagamentos).

    Um mesmo Pagamento nunca aparece em dois resultados EXATO_1X1/EXATO_SOMA
    diferentes — o pool é compartilhado entre todos os SISPAG e mutado a
    cada match bem-sucedido.
    """
    pool_por_data: dict[date, list[Pagamento]] = {}
    for pagamento in pagamentos:
        pool_por_data.setdefault(pagamento.data_efetiva, []).append(pagamento)

    resultados: list[ResultadoSispag] = []
    for sequencia, data_sispag, valor_sispag in sispags:
        pool_da_data = pool_por_data.setdefault(data_sispag, [])
        resultado = _conciliar_um_sispag(sequencia, data_sispag, valor_sispag, pool_da_data)
        resultados.append(resultado)

    return resultados
