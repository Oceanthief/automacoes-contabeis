"""Validação contábil final: determina se cada linha de saida_contabil.xlsx
(já com as correções manuais aplicadas) está apta a virar TXT, e se o
arquivo inteiro está APROVADO_PARA_TXT ou BLOQUEADO_PARA_TXT.

Só valida — nunca corrige, nunca inventa conta, nunca altera dado. Cada
verificação devolve uma lista de motivos de bloqueio (vazia = passou).
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Optional

import pandas as pd

from enriquecimento.classificacao_extrato import CONTA_BANCO, CONTA_TARIFA
from normalizacao import normalizar_conta_contabil, valor_para_decimal

STATUS_APROVADO = "APROVADO"
STATUS_BLOQUEADO = "BLOQUEADO"

STATUS_GERAL_APROVADO = "APROVADO_PARA_TXT"
STATUS_GERAL_BLOQUEADO = "BLOQUEADO_PARA_TXT"

STATUS_DE_PARA_BLOQUEANTES = {"SEM_DE_PARA", "AMBIGUO"}

DESCRICAO_SISPAG_FORNECEDORES = "SISPAG FORNECEDORES"
PREFIXO_TARIFA = "TAR "

TEXTOS_DESCRICAO_INVALIDOS = {"", "NAN", "NONE", "SEM INFORMACAO"}


# ---------------------------------------------------------------------------
# Validações individuais — cada uma devolve uma lista de motivos (vazia = OK)
# ---------------------------------------------------------------------------

def _validar_data(valor, competencia: Optional[tuple]) -> list[str]:
    if pd.isna(valor):
        return ["DATA vazia"]
    try:
        data = pd.Timestamp(valor)
    except (ValueError, TypeError):
        return ["DATA inválida (não convertível)"]

    if competencia is not None:
        inicio, fim = competencia
        if not (inicio <= data <= fim):
            return [f"DATA fora da competência processada ({data.date()} fora de {inicio.date()}–{fim.date()})"]
    return []


def _validar_valor(valor) -> list[str]:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ["VALOR vazio"]
    try:
        decimal_valor = valor_para_decimal(valor)
    except (InvalidOperation, ValueError, TypeError):
        return ["VALOR não numérico/inválido"]
    if decimal_valor == 0:
        return ["VALOR igual a zero"]
    if decimal_valor < 0:
        return ["VALOR negativo"]
    return []


def _validar_debito_credito_preenchidos(debito, credito) -> list[str]:
    motivos = []
    if not normalizar_conta_contabil(debito):
        motivos.append("DEBITO vazio")
    if not normalizar_conta_contabil(credito):
        motivos.append("CREDITO vazio")
    return motivos


def _validar_conta_bancaria(tipo_movimento: str, debito, credito) -> list[str]:
    debito_norm = normalizar_conta_contabil(debito)
    credito_norm = normalizar_conta_contabil(credito)
    if tipo_movimento == "PAGAMENTO" and credito_norm and credito_norm != CONTA_BANCO:
        return [f"PAGAMENTO com CREDITO diferente de {CONTA_BANCO} (veio {credito_norm})"]
    if tipo_movimento == "RECEBIMENTO" and debito_norm and debito_norm != CONTA_BANCO:
        return [f"RECEBIMENTO com DEBITO diferente de {CONTA_BANCO} (veio {debito_norm})"]
    return []


def _validar_contas_iguais(debito, credito) -> list[str]:
    debito_norm = normalizar_conta_contabil(debito)
    credito_norm = normalizar_conta_contabil(credito)
    if debito_norm and credito_norm and debito_norm == credito_norm:
        return ["ERRO_CONTAS_IGUAIS: DEBITO == CREDITO"]
    return []


def _validar_de_para(status_de_para: str) -> list[str]:
    status = str(status_de_para).strip()
    if status in STATUS_DE_PARA_BLOQUEANTES:
        return [f"DE-PARA não resolvido ({status})"]
    return []


def _validar_regra_fixa_tarifa(descricao_original: str, tipo_movimento: str, debito, credito) -> list[str]:
    descricao = str(descricao_original).strip()
    if tipo_movimento == "PAGAMENTO" and descricao.startswith(PREFIXO_TARIFA):
        debito_norm = normalizar_conta_contabil(debito)
        credito_norm = normalizar_conta_contabil(credito)
        motivos = []
        if debito_norm != CONTA_TARIFA:
            motivos.append(f"Tarifa de pagamento deveria ter DEBITO={CONTA_TARIFA} (veio {debito_norm})")
        if credito_norm != CONTA_BANCO:
            motivos.append(f"Tarifa de pagamento deveria ter CREDITO={CONTA_BANCO} (veio {credito_norm})")
        return motivos
    return []


def _validar_descricao(nova_descricao) -> list[str]:
    texto = "" if pd.isna(nova_descricao) else str(nova_descricao).strip()
    if texto.upper() in TEXTOS_DESCRICAO_INVALIDOS:
        return ["NOVA_DESCRICAO vazia ou apenas espaços"]
    return []


# ---------------------------------------------------------------------------
# Competência: derivada dos próprios dados (moda do mês), nunca do nome do arquivo
# ---------------------------------------------------------------------------

def determinar_competencia(df: pd.DataFrame) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Deriva o intervalo de competência a partir do mês mais frequente entre
    as DATA da própria base — nunca do nome do arquivo. Uma linha isolada em
    outro mês fica fora do intervalo e é reportada como suspeita."""
    datas_validas = pd.to_datetime(df["DATA"], errors="coerce").dropna()
    if datas_validas.empty:
        return None
    meses = datas_validas.dt.to_period("M")
    mes_mais_comum = Counter(meses).most_common(1)[0][0]
    inicio = mes_mais_comum.start_time
    fim = mes_mais_comum.end_time
    return (inicio, fim)


# ---------------------------------------------------------------------------
# Validações que dependem de mais de uma linha (duplicidade, SISPAG)
# ---------------------------------------------------------------------------

def _sequencias_bancarias_problematicas(auditoria_granularizacao: pd.DataFrame) -> dict[int, str]:
    """Mapa SEQUENCIA_BANCO -> motivo, para SISPAG que a Fase 5 não fechou."""
    problematicos: dict[int, str] = {}
    for _, linha in auditoria_granularizacao.iterrows():
        status_gran = str(linha["STATUS_GRANULARIZACAO"]).strip()
        if status_gran != "OK":
            problematicos[int(linha["SEQUENCIA_BANCO"])] = (
                f"SISPAG {status_gran} (STATUS_CONCILIACAO={linha['STATUS_CONCILIACAO']})"
            )
    return problematicos


def _titulos_duplicados(df: pd.DataFrame) -> set[str]:
    titulos = df["TITULO"].dropna().astype(str)
    titulos = titulos[titulos != ""]
    contagem = titulos.value_counts()
    return set(contagem[contagem > 1].index)


# ---------------------------------------------------------------------------
# Validação completa
# ---------------------------------------------------------------------------

def validar_saida(
    saida_df: pd.DataFrame, auditoria_granularizacao_df: pd.DataFrame
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Adiciona STATUS_VALIDACAO e MOTIVO_BLOQUEIO a cada linha de saida_df.

    Devolve (cópia_validada, contagem_por_categoria) — nunca altera saida_df
    no lugar. contagem_por_categoria mostra quantas linhas falharam em cada
    uma das checagens (uma linha pode falhar em mais de uma).
    """
    resultado = saida_df.copy()

    if not resultado["SEQUENCIA_CONTABIL"].is_unique:
        contagem = resultado["SEQUENCIA_CONTABIL"].value_counts()
        duplicadas = set(contagem[contagem > 1].index)
        raise ValueError(f"SEQUENCIA_CONTABIL duplicada (dado inválido, não deveria acontecer): {duplicadas}")

    competencia = determinar_competencia(resultado)
    sequencias_problematicas = _sequencias_bancarias_problematicas(auditoria_granularizacao_df)
    titulos_duplicados = _titulos_duplicados(resultado)

    todos_motivos: list[list[str]] = []
    contagem_categoria: dict[str, int] = {
        "DATA": 0, "VALOR": 0, "DEBITO_CREDITO_VAZIO": 0, "CONTA_BANCARIA": 0,
        "CONTAS_IGUAIS": 0, "DE_PARA": 0, "REGRA_FIXA_TARIFA": 0, "DESCRICAO": 0,
        "SISPAG": 0, "DUPLICIDADE": 0,
    }

    for _, linha in resultado.iterrows():
        checagens = {
            "DATA": _validar_data(linha["DATA"], competencia),
            "VALOR": _validar_valor(linha["VALOR"]),
            "DEBITO_CREDITO_VAZIO": _validar_debito_credito_preenchidos(linha["DEBITO"], linha["CREDITO"]),
            "CONTA_BANCARIA": _validar_conta_bancaria(linha["TIPO_MOVIMENTO"], linha["DEBITO"], linha["CREDITO"]),
            "CONTAS_IGUAIS": _validar_contas_iguais(linha["DEBITO"], linha["CREDITO"]),
            "DE_PARA": _validar_de_para(linha["STATUS_DE_PARA"]),
            "REGRA_FIXA_TARIFA": _validar_regra_fixa_tarifa(
                linha["DESCRICAO_ORIGINAL"], linha["TIPO_MOVIMENTO"], linha["DEBITO"], linha["CREDITO"]
            ),
            "DESCRICAO": _validar_descricao(linha["NOVA_DESCRICAO"]),
            "SISPAG": [],
            "DUPLICIDADE": [],
        }

        seq_banco = int(linha["SEQUENCIA_BANCO"])
        if seq_banco in sequencias_problematicas:
            checagens["SISPAG"].append(sequencias_problematicas[seq_banco])

        titulo = str(linha.get("TITULO", "")).strip()
        if titulo and titulo in titulos_duplicados:
            checagens["DUPLICIDADE"].append(f"PAGAMENTO_DUPLICADO: TITULO {titulo} aparece mais de uma vez")

        motivos_linha: list[str] = []
        for categoria, motivos in checagens.items():
            if motivos:
                contagem_categoria[categoria] += 1
                motivos_linha.extend(motivos)

        todos_motivos.append(motivos_linha)

    resultado["STATUS_VALIDACAO"] = [STATUS_BLOQUEADO if m else STATUS_APROVADO for m in todos_motivos]
    resultado["MOTIVO_BLOQUEIO"] = ["; ".join(m) for m in todos_motivos]

    return resultado, contagem_categoria


def status_geral(saida_validada: pd.DataFrame) -> str:
    if (saida_validada["STATUS_VALIDACAO"] == STATUS_BLOQUEADO).any():
        return STATUS_GERAL_BLOQUEADO
    return STATUS_GERAL_APROVADO


# ---------------------------------------------------------------------------
# Totalização — granularização aumenta linhas, nunca aumenta valor econômico
# ---------------------------------------------------------------------------

@dataclass
class Totais:
    qtd_movimentos_bancarios: int
    qtd_linhas_contabeis: int
    total_pagamentos: Decimal
    total_recebimentos: Decimal
    total_sispag_granularizado: Decimal
    total_bancario_sispag: Decimal
    total_debitos: Decimal
    total_creditos: Decimal


def calcular_totais(saida_df: pd.DataFrame, auditoria_granularizacao_df: Optional[pd.DataFrame] = None) -> Totais:
    """Calcula os totais pedidos pela Validação 12.

    total_bancario_sispag vem de auditoria_granularizacao_df (coluna
    VALOR_BANCO), NUNCA reconstruído a partir das linhas já granularizadas —
    cada linha ali só carrega o valor do pagamento individual, então somar
    "a primeira linha de cada SEQUENCIA_BANCO" dá o valor errado (bug real
    pego pelo teste desta função: uma primeira versão fazia exatamente isso).
    """
    valores = saida_df["VALOR"].apply(valor_para_decimal)

    pagamentos = saida_df["TIPO_MOVIMENTO"] == "PAGAMENTO"
    recebimentos = saida_df["TIPO_MOVIMENTO"] == "RECEBIMENTO"
    sispag = saida_df["DESCRICAO_ORIGINAL"] == DESCRICAO_SISPAG_FORNECEDORES

    if auditoria_granularizacao_df is not None and len(auditoria_granularizacao_df):
        total_bancario_sispag = auditoria_granularizacao_df["VALOR_BANCO"].apply(valor_para_decimal).sum()
    else:
        total_bancario_sispag = Decimal("0.00")

    return Totais(
        qtd_movimentos_bancarios=saida_df["SEQUENCIA_BANCO"].nunique(),
        qtd_linhas_contabeis=len(saida_df),
        total_pagamentos=valores[pagamentos].sum() if pagamentos.any() else Decimal("0.00"),
        total_recebimentos=valores[recebimentos].sum() if recebimentos.any() else Decimal("0.00"),
        total_sispag_granularizado=valores[sispag].sum() if sispag.any() else Decimal("0.00"),
        total_bancario_sispag=total_bancario_sispag,
        total_debitos=valores.sum(),  # cada linha é um lançamento partida-dobrada: débito=crédito=VALOR
        total_creditos=valores.sum(),
    )
