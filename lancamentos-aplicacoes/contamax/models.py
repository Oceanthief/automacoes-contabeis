"""
Estruturas de dados usadas em todo o pipeline CONTAMAX.

Usar dataclasses aqui (em vez de dicts soltos) dá duas vantagens práticas:
1. o editor/typechecker avisa se algum campo for esquecido ou escrito errado;
2. fica explícito, num único lugar, qual é o "contrato" de dados entre as
   camadas de leitura, regra de negócio e (nas próximas etapas) exportação.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from enum import Enum


class StatusLancamento(str, Enum):
    """
    Resultado do processamento de uma data específica.

    Existe um status por data (não por linha do extrato) porque a regra
    contábil do CONTAMAX opera sobre o total agregado do dia, não sobre
    movimentações individuais.
    """

    GERADO = "GERADO"
    # Débito == crédito no dia: decisão 5.1 da especificação é não gerar
    # lançamento, mas o dia ainda aparece no resultado para dar rastreabilidade
    # (auditoria) de que ele foi processado e por que não gerou lançamento.
    SEM_LANCAMENTO_EMPATE = "SEM_LANCAMENTO_EMPATE"
    # Alguma linha do extrato daquele dia tinha um valor monetário que não
    # pôde ser convertido com segurança (ver contamax.valores). O dia inteiro
    # é marcado como erro em vez de somar parcialmente — somar "o que deu
    # certo" e ignorar o resto geraria um total contábil errado sem avisar
    # ninguém, que é exatamente o tipo de bug perigoso em automação contábil.
    ERRO_VALOR_INVALIDO = "ERRO_VALOR_INVALIDO"


@dataclass(frozen=True)
class Movimentacao:
    """Uma linha (já normalizada) do extrato bancário."""

    data: date
    debito: Decimal
    credito: Decimal
    historico: str = ""


@dataclass(frozen=True)
class LinhaComErro:
    """
    Uma linha do extrato cujo valor monetário não pôde ser convertido com
    segurança (ver contamax.valores.extrair_valor_monetario).

    `data` pode ser None no caso raro de a própria célula de data também
    estar inválida — nesse caso a linha não pode nem ser associada a um dia,
    e a leitura deve interromper o processamento (ver contamax.leitura_xls).
    """

    numero_linha: int  # linha da planilha (1-indexado, para o usuário conferir)
    data: date | None
    mensagem: str


@dataclass(frozen=True)
class ResumoExecucao:
    """Números de fechamento de uma rodada de processamento, para o terminal."""

    periodo: str  # "AAAA-MM"
    dias_processados: int
    total_aplicacoes: int
    total_resgates: int
    pendencias: int


@dataclass(frozen=True)
class LancamentoContamax:
    """
    Um resultado por data processada — é exatamente o esquema pedido nos
    critérios de aceite (DATA, TOTAL_DEBITO, ... STATUS).

    Para dias com STATUS != GERADO, conta_debito/conta_credito/historico/valor
    ficam como None: não existe lançamento contábil para aquele dia, então
    não faz sentido preencher esses campos com um valor "de mentira" (como 0
    ou string vazia) que alguém possa confundir com um lançamento real.
    """

    data: date
    total_debito: Decimal
    total_credito: Decimal
    diferenca: Decimal
    status: StatusLancamento
    conta_debito: int | None = None
    conta_credito: int | None = None
    historico: str | None = None
    valor: Decimal | None = None
