"""Motor de DE-PARA: encontra a conta contábil para um TEXTO_MATCH usando as
regras da planilha (leitura/de_para.py).

Genérico de propósito: não sabe nada sobre extrato, SISPAG ou pagamentos
individuais — recebe (tipo_movimento, texto_match_normalizado) e devolve um
resultado. Isso permite reaproveitar exatamente o mesmo motor tanto para os
lançamentos comuns do extrato quanto, na próxima fase, para cada pagamento
individual granularizado de um SISPAG.

Não faz fuzzy matching — só EXATO, COMECA_COM (prefixo) e CONTEM (substring),
sempre sobre texto já normalizado (normalizacao.normalizar_texto_classificacao).

ORDEM DE RESOLUÇÃO
Entre as regras ATIVAS e compatíveis com o TIPO_MOVIMENTO que casam
textualmente, vence a de menor PRIORIDADE numérica; em caso de empate,
desempata por especificidade do TIPO_MATCH (EXATO < COMECA_COM < CONTEM).
Se ainda houver 2+ regras empatadas apontando para CONTA_CONTABIL
diferentes, o resultado é AMBIGUO — nenhuma conta é atribuída. Regras
empatadas que apontam para a MESMA conta não são tratadas como ambiguidade
(é redundância de cadastro, não conflito de decisão).
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

import pandas as pd

from normalizacao import normalizar_conta_contabil, normalizar_texto_classificacao

STATUS_MATCH_EXATO = "MATCH_EXATO"
STATUS_MATCH_REGRA = "MATCH_REGRA"
STATUS_AMBIGUO = "AMBIGUO"
STATUS_SEM_DE_PARA = "SEM_DE_PARA"
STATUS_JA_CLASSIFICADO = "JA_CLASSIFICADO"

TIPOS_MOVIMENTO_VALIDOS = {"PAGAMENTO", "RECEBIMENTO", "AMBOS"}
TIPOS_MATCH_VALIDOS = {"EXATO", "COMECA_COM", "CONTEM"}
_ESPECIFICIDADE = {"EXATO": 0, "COMECA_COM": 1, "CONTEM": 2}


class MotorDeParaError(Exception):
    """Erro de integridade de uma regra do DE-PARA (linha malformada)."""


@dataclass
class RegraDePara:
    """Uma regra do DE-PARA já validada e normalizada."""

    linha_origem: int  # número da linha no Excel, só para rastreabilidade
    chave_match: str
    chave_match_normalizada: str
    tipo_movimento: str
    tipo_match: str
    conta_contabil: str
    prioridade: int
    ativo: bool
    historico_padrao: str
    observacao: str


@dataclass
class ResultadoDePara:
    """Resultado da tentativa de classificar um TEXTO_MATCH."""

    status: str
    conta_encontrada: Optional[str] = None
    tipo_match_utilizado: Optional[str] = None
    chave_match_utilizada: Optional[str] = None
    prioridade_utilizada: Optional[int] = None
    historico_padrao_utilizado: Optional[str] = None
    observacao: str = ""


def carregar_regras(df: pd.DataFrame) -> list[RegraDePara]:
    """Converte o DataFrame do DE-PARA em RegraDePara, validando cada linha.

    Levanta MotorDeParaError apontando a linha do Excel se algo estiver fora
    do esperado — nunca ignora silenciosamente uma linha malformada, e nunca
    aceita uma regra ATIVA com conta placeholder/vazia (ex.: as linhas de
    exemplo "<preencher>" do modelo, que já vêm com ATIVO=NAO de propósito).
    """
    regras: list[RegraDePara] = []
    for indice, linha in df.iterrows():
        numero_linha_excel = indice + 2  # +1 cabeçalho, +1 base 1 do Excel

        tipo_movimento = str(linha["TIPO_MOVIMENTO"]).strip().upper()
        if tipo_movimento not in TIPOS_MOVIMENTO_VALIDOS:
            raise MotorDeParaError(
                f"Linha {numero_linha_excel}: TIPO_MOVIMENTO inválido ({linha['TIPO_MOVIMENTO']!r})."
            )

        tipo_match = str(linha["TIPO_MATCH"]).strip().upper()
        if tipo_match not in TIPOS_MATCH_VALIDOS:
            raise MotorDeParaError(
                f"Linha {numero_linha_excel}: TIPO_MATCH inválido ({linha['TIPO_MATCH']!r})."
            )

        ativo_raw = str(linha["ATIVO"]).strip().upper()
        if ativo_raw not in ("SIM", "NAO"):
            raise MotorDeParaError(f"Linha {numero_linha_excel}: ATIVO inválido ({linha['ATIVO']!r}).")
        ativo = ativo_raw == "SIM"

        if pd.isna(linha["PRIORIDADE"]):
            raise MotorDeParaError(f"Linha {numero_linha_excel}: PRIORIDADE ausente.")
        prioridade = int(linha["PRIORIDADE"])

        conta_contabil = normalizar_conta_contabil(linha["CONTA_CONTABIL"])
        if ativo and (not conta_contabil or conta_contabil.strip().startswith("<")):
            raise MotorDeParaError(
                f"Linha {numero_linha_excel}: regra ATIVA com CONTA_CONTABIL inválida/placeholder "
                f"({linha['CONTA_CONTABIL']!r}). Corrija o DE-PARA ou marque ATIVO=NAO."
            )

        chave_match = str(linha["CHAVE_MATCH"]).strip()
        historico_padrao = "" if pd.isna(linha.get("HISTORICO_PADRAO")) else str(linha["HISTORICO_PADRAO"])
        observacao = "" if pd.isna(linha.get("OBSERVACAO")) else str(linha["OBSERVACAO"])

        regras.append(
            RegraDePara(
                linha_origem=numero_linha_excel,
                chave_match=chave_match,
                chave_match_normalizada=normalizar_texto_classificacao(chave_match),
                tipo_movimento=tipo_movimento,
                tipo_match=tipo_match,
                conta_contabil=conta_contabil,
                prioridade=prioridade,
                ativo=ativo,
                historico_padrao=historico_padrao,
                observacao=observacao,
            )
        )
    return regras


def _candidatos(regras: list[RegraDePara], tipo_movimento: str, texto_match: str) -> list[RegraDePara]:
    candidatos = []
    for regra in regras:
        if not regra.ativo:
            continue
        if regra.tipo_movimento not in (tipo_movimento, "AMBOS"):
            continue
        if regra.tipo_match == "EXATO" and regra.chave_match_normalizada == texto_match:
            candidatos.append(regra)
        elif (
            regra.tipo_match == "COMECA_COM"
            and regra.chave_match_normalizada
            and texto_match.startswith(regra.chave_match_normalizada)
        ):
            candidatos.append(regra)
        elif (
            regra.tipo_match == "CONTEM"
            and regra.chave_match_normalizada
            and regra.chave_match_normalizada in texto_match
        ):
            candidatos.append(regra)
    return candidatos


def classificar(regras: list[RegraDePara], tipo_movimento: str, texto_match: str) -> ResultadoDePara:
    """Encontra a(s) regra(s) vencedora(s) para um TEXTO_MATCH já normalizado."""
    if not texto_match:
        return ResultadoDePara(STATUS_SEM_DE_PARA, observacao="TEXTO_MATCH vazio.")

    candidatos = _candidatos(regras, tipo_movimento, texto_match)
    if not candidatos:
        return ResultadoDePara(STATUS_SEM_DE_PARA, observacao="Nenhuma regra ativa corresponde.")

    candidatos.sort(key=lambda r: (r.prioridade, _ESPECIFICIDADE[r.tipo_match]))
    melhor_chave = (candidatos[0].prioridade, _ESPECIFICIDADE[candidatos[0].tipo_match])
    vencedores = [r for r in candidatos if (r.prioridade, _ESPECIFICIDADE[r.tipo_match]) == melhor_chave]

    contas_distintas = {v.conta_contabil for v in vencedores}
    if len(contas_distintas) > 1:
        detalhe = "; ".join(f"{v.chave_match} -> {v.conta_contabil}" for v in vencedores)
        return ResultadoDePara(
            STATUS_AMBIGUO,
            observacao=f"{len(vencedores)} regras empatadas (prioridade {melhor_chave[0]}) apontam contas diferentes: {detalhe}",
        )

    escolhida = vencedores[0]
    status = STATUS_MATCH_EXATO if escolhida.tipo_match == "EXATO" else STATUS_MATCH_REGRA
    observacao = ""
    if len(vencedores) > 1:
        observacao = f"{len(vencedores)} regras empatadas resolveram para a mesma conta ({escolhida.conta_contabil})."

    return ResultadoDePara(
        status,
        conta_encontrada=escolhida.conta_contabil,
        tipo_match_utilizado=escolhida.tipo_match,
        chave_match_utilizada=escolhida.chave_match,
        prioridade_utilizada=escolhida.prioridade,
        historico_padrao_utilizado=escolhida.historico_padrao or None,
        observacao=observacao,
    )
