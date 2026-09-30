"""Granularização dos SISPAG FORNECEDORES conciliados em linhas contábeis
individuais — uma por pagamento (por TITULO), nunca uma por VALOR_CLASS.

Reaproveita:
  - granularizacao.reconstrucao_sispag para os pagamentos estruturados;
  - enriquecimento.motor_de_para para classificar cada pagamento individual
    usando o NOME do fornecedor como fonte de TEXTO_MATCH.

Não reimplementa nenhuma das camadas anteriores.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Optional

from conciliacao.motor_sispag import STATUS_EXATO_1X1, STATUS_EXATO_SOMA, ResultadoSispag
from enriquecimento.motor_de_para import RegraDePara, classificar as classificar_de_para
from normalizacao import normalizar_texto_classificacao, para_centavos

DESCRICAO_SISPAG_FORNECEDORES = "SISPAG FORNECEDORES"
CONTA_BANCO = "10001"

TIPO_ORIGEM_NORMAL = "NORMAL"
TIPO_ORIGEM_SISPAG_1X1 = "SISPAG_1X1"
TIPO_ORIGEM_SISPAG_GRANULARIZADO = "SISPAG_GRANULARIZADO"

STATUS_GRAN_OK = "OK"
STATUS_GRAN_ERRO = "ERRO_GRANULARIZACAO"
# SISPAG cuja conciliação não chegou a EXATO_1X1/EXATO_SOMA (AMBIGUO ou
# NAO_ENCONTRADO na Fase de conciliação): não há pagamento nenhum para
# granularizar. Não é um erro desta fase — é herdado de uma fase anterior.
STATUS_GRAN_NAO_CONCILIADO = "SISPAG_NAO_CONCILIADO"


class GranularizacaoError(Exception):
    """Erro bloqueante de integridade (ex.: PAGAMENTO_DUPLICADO)."""


@dataclass
class LinhaContabil:
    sequencia_contabil: int
    sequencia_banco: int
    debito: str
    credito: str
    data: object
    valor: Decimal
    descricao_original: str
    nova_descricao: str
    tipo_movimento: str
    tipo_origem: str
    status_conciliacao: str
    status_de_para: str
    titulo: str
    fornecedor: str
    nf_num: str
    historico: str
    forma_pgto: str
    conta_origem: str


@dataclass
class AuditoriaSispagBanco:
    sequencia_banco: int
    data: object
    valor_banco: Decimal
    status_conciliacao: str
    qtd_pagamentos_esperados: int
    qtd_linhas_geradas: int
    total_granularizado: Decimal
    diferenca: Decimal
    status_granularizacao: str
    titulos: str
    fornecedores: str
    observacao: str


def granularizar_sispag(
    resultados_sispag: list[ResultadoSispag],
    regras_de_para: list[RegraDePara],
) -> tuple[list[LinhaContabil], list[AuditoriaSispagBanco], list[str]]:
    """Expande cada SISPAG conciliado (EXATO_1X1/EXATO_SOMA) em 1+ linhas
    contábeis — uma por pagamento (TITULO) já encontrado pela conciliação.

    Retorna (linhas_contabeis, auditoria_por_sispag_banco,
    casos_requer_granularizacao_classificacao). Levanta GranularizacaoError
    (PAGAMENTO_DUPLICADO) se o mesmo pagamento aparecer em dois SISPAG
    diferentes — nunca deveria acontecer, dado que o motor de conciliação já
    garante isso, mas é conferido aqui de novo, explicitamente.
    """
    linhas: list[LinhaContabil] = []
    auditorias: list[AuditoriaSispagBanco] = []
    requer_granularizacao_classificacao: list[str] = []

    pagamentos_ja_usados: set[tuple] = set()

    for resultado in resultados_sispag:
        if resultado.status not in (STATUS_EXATO_1X1, STATUS_EXATO_SOMA):
            auditorias.append(
                AuditoriaSispagBanco(
                    sequencia_banco=resultado.sequencia,
                    data=resultado.data,
                    valor_banco=resultado.valor_sispag,
                    status_conciliacao=resultado.status,
                    qtd_pagamentos_esperados=0,
                    qtd_linhas_geradas=0,
                    total_granularizado=Decimal("0.00"),
                    diferenca=resultado.valor_sispag,
                    status_granularizacao=STATUS_GRAN_NAO_CONCILIADO,
                    titulos="",
                    fornecedores="",
                    observacao=(
                        "SISPAG não chegou a EXATO_1X1/EXATO_SOMA na conciliação "
                        f"({resultado.status}) — nada para granularizar nesta fase."
                    ),
                )
            )
            continue

        pagamentos_ordenados = sorted(resultado.pagamentos, key=lambda p: p.titulo)
        tipo_origem = (
            TIPO_ORIGEM_SISPAG_1X1 if len(pagamentos_ordenados) == 1 else TIPO_ORIGEM_SISPAG_GRANULARIZADO
        )

        total_granularizado_centavos = 0

        for pagamento in pagamentos_ordenados:
            chave_pagamento = (pagamento.titulo, pagamento.data_efetiva)
            if chave_pagamento in pagamentos_ja_usados:
                raise GranularizacaoError(
                    f"PAGAMENTO_DUPLICADO: TITULO {pagamento.titulo} ({pagamento.data_efetiva}) "
                    f"já foi granularizado em outro SISPAG antes de SEQUENCIA_BANCO {resultado.sequencia}."
                )
            pagamentos_ja_usados.add(chave_pagamento)

            if len(pagamento.classificacoes) > 1:
                requer_granularizacao_classificacao.append(
                    f"TITULO {pagamento.titulo} (SEQUENCIA_BANCO {resultado.sequencia}, "
                    f"fornecedor {pagamento.nome}): {len(pagamento.classificacoes)} classificações "
                    "somadas em soma(VALOR_CLASS) — não divididas nesta fase."
                )

            texto_match = normalizar_texto_classificacao(pagamento.nome)
            resultado_de_para = classificar_de_para(regras_de_para, "PAGAMENTO", texto_match)
            debito = resultado_de_para.conta_encontrada or ""

            linhas.append(
                LinhaContabil(
                    sequencia_contabil=0,  # numerada no final, após juntar com as linhas NORMAL
                    sequencia_banco=resultado.sequencia,
                    debito=debito,
                    credito=CONTA_BANCO,
                    data=resultado.data,
                    valor=pagamento.valor_classificado,
                    descricao_original=DESCRICAO_SISPAG_FORNECEDORES,
                    nova_descricao=f"VLR REF. PGTO {pagamento.nome}",
                    tipo_movimento="PAGAMENTO",
                    tipo_origem=tipo_origem,
                    status_conciliacao=resultado.status,
                    status_de_para=resultado_de_para.status,
                    titulo=str(pagamento.titulo),
                    fornecedor=pagamento.nome,
                    nf_num=pagamento.nf_num or "",
                    historico=pagamento.historico,
                    forma_pgto=pagamento.forma_pgto,
                    conta_origem=pagamento.conta,
                )
            )
            total_granularizado_centavos += para_centavos(pagamento.valor_classificado)

        valor_banco_centavos = para_centavos(resultado.valor_sispag)
        diferenca_centavos = valor_banco_centavos - total_granularizado_centavos
        status_granularizacao = STATUS_GRAN_OK if diferenca_centavos == 0 else STATUS_GRAN_ERRO

        auditorias.append(
            AuditoriaSispagBanco(
                sequencia_banco=resultado.sequencia,
                data=resultado.data,
                valor_banco=resultado.valor_sispag,
                status_conciliacao=resultado.status,
                qtd_pagamentos_esperados=len(pagamentos_ordenados),
                qtd_linhas_geradas=len(pagamentos_ordenados),
                total_granularizado=Decimal(total_granularizado_centavos) / 100,
                diferenca=Decimal(diferenca_centavos) / 100,
                status_granularizacao=status_granularizacao,
                titulos="; ".join(str(p.titulo) for p in pagamentos_ordenados),
                fornecedores="; ".join(p.nome for p in pagamentos_ordenados),
                observacao=(
                    ""
                    if status_granularizacao == STATUS_GRAN_OK
                    else "Soma granularizada não fecha com o valor bancário em centavos."
                ),
            )
        )

    return linhas, auditorias, requer_granularizacao_classificacao
