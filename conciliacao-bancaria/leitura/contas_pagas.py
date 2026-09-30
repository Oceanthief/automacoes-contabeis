"""Preparação e normalização do arquivo de Contas Pagas.

A Fase 0 confirmou: um mesmo TITULO pode ter várias linhas porque o
pagamento tem várias classificações contábeis (VALOR_CLASS). Este módulo
reconstrói cada pagamento físico como uma única entidade — nunca uma linha
por classificação.

Fora do escopo aqui: qualquer leitura do Extrato ou tentativa de casar isso
com SISPAG FORNECEDORES (isso é a próxima fase).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional

import pandas as pd

from config import CONTA_BANCARIA_ALVO
from normalizacao import normalizar_data, valor_para_decimal

COLUNAS_ESPERADAS = [
    "TITULO",
    "NOME",
    "NF_NUM",
    "DATA_EFETIVA_CH",
    "VALOR",
    "VALOR_CLASS",
    "CLASSIFIC_COD",
    "CLASSIFIC_DESC",
    "HISTORICO",
    "FORMA_PGTO",
    "CONTA",
]

# Conta usada nesta automação bancária (definida em config.py / variável de
# ambiente). Outras contas (ex.: "341 - Cartão Corporativo") ficam fora da
# base de candidatos, mas continuam disponíveis no relatório — decisão
# fechada na Fase 0.
CONTA_ALVO = CONTA_BANCARIA_ALVO

TOLERANCIA_DIVERGENCIA = Decimal("0.01")

# Campos que, por definição, têm que ser os mesmos em todas as linhas de
# classificação de um mesmo pagamento. Se variarem, a premissa de "é um
# único pagamento físico" está quebrada — melhor travar alto e claro do que
# silenciosamente escolher um valor.
CAMPOS_INVARIANTES = ["NOME", "CONTA", "FORMA_PGTO"]


class ContasPagasError(Exception):
    """Erros de leitura, formato ou integridade do arquivo de Contas Pagas."""


@dataclass
class Classificacao:
    """Uma linha de classificação contábil dentro de um pagamento."""

    codigo: str
    descricao: str
    valor_class: Decimal


@dataclass
class Pagamento:
    """Um pagamento físico reconstruído a partir de 1+ linhas de classificação.

    A chave de identidade é (titulo, data_efetiva) — nunca titulo sozinho,
    conforme pedido: um mesmo número de título poderia em tese se repetir em
    datas diferentes (não ocorre na amostra de agosto/2026, mas o código não
    deve depender disso).
    """

    titulo: int
    data_efetiva: date
    nome: str
    nf_num: Optional[str]
    forma_pgto: str
    conta: str
    historico: str
    valor_original: Optional[Decimal]
    valor_classificado: Decimal
    classificacoes: list[Classificacao] = field(default_factory=list)

    @property
    def divergencia(self) -> Optional[Decimal]:
        """VALOR original menos a soma de VALOR_CLASS. None se VALOR ausente."""
        if self.valor_original is None:
            return None
        return self.valor_original - self.valor_classificado

    @property
    def tem_divergencia_relevante(self) -> bool:
        div = self.divergencia
        return div is not None and abs(div) > TOLERANCIA_DIVERGENCIA


def ler_contas_pagas(caminho: str | Path) -> pd.DataFrame:
    """Lê o arquivo de Contas Pagas e valida a presença das colunas esperadas."""
    caminho = Path(caminho)
    if not caminho.exists():
        raise ContasPagasError(f"Arquivo não encontrado: {caminho}")

    try:
        df = pd.read_excel(caminho, sheet_name=0)
    except Exception as exc:
        raise ContasPagasError(f"Não foi possível abrir '{caminho}': {exc}") from exc

    faltantes = [c for c in COLUNAS_ESPERADAS if c not in df.columns]
    if faltantes:
        raise ContasPagasError(f"Colunas ausentes em Contas Pagas: {faltantes}")

    return df


def _valor_unico_nao_nulo(serie: pd.Series, titulo, data_efetiva) -> Optional[Decimal]:
    """Pega o VALOR do grupo, exigindo que, se existir, seja o mesmo em todas as linhas."""
    valores = serie.dropna().unique()
    if len(valores) == 0:
        return None
    if len(valores) > 1:
        raise ContasPagasError(
            f"TITULO {titulo} ({data_efetiva}) tem mais de um VALOR distinto "
            f"e não-nulo entre suas linhas de classificação: {list(valores)}"
        )
    return valor_para_decimal(valores[0])


def _campo_invariante(serie: pd.Series, campo: str, titulo, data_efetiva) -> str:
    """Pega um campo que deveria ser idêntico em todas as linhas do grupo.

    Levanta erro se variar — ver CAMPOS_INVARIANTES para o porquê.
    """
    valores = serie.dropna().unique()
    if len(valores) == 0:
        return ""
    if len(valores) > 1:
        raise ContasPagasError(
            f"TITULO {titulo} ({data_efetiva}): campo '{campo}' varia entre linhas "
            f"de classificação do mesmo pagamento: {list(valores)}. "
            "Isso quebra a premissa de que é um único pagamento físico."
        )
    return str(valores[0])


def _campo_preferencial(serie: pd.Series) -> str:
    """Pega o primeiro valor não-nulo de um campo que pode variar sem ser fatal
    (ex.: HISTORICO, NF_NUM às vezes trazem observações por classificação)."""
    valores = serie.dropna()
    return str(valores.iloc[0]) if len(valores) else ""


def montar_pagamentos(df: pd.DataFrame) -> tuple[list[Pagamento], list[Pagamento]]:
    """Agrupa por (TITULO, DATA_EFETIVA_CH) e reconstrói pagamentos únicos.

    Retorna uma tupla (candidatos, fora_do_escopo):
      - candidatos: pagamentos com CONTA == CONTA_ALVO (base para a conciliação);
      - fora_do_escopo: pagamentos de outras contas — preservados só para relatório.
    """
    trabalho = df.copy()
    trabalho["_data_normalizada"] = trabalho["DATA_EFETIVA_CH"].apply(normalizar_data)

    candidatos: list[Pagamento] = []
    fora_do_escopo: list[Pagamento] = []

    for (titulo, data_efetiva), grupo in trabalho.groupby(["TITULO", "_data_normalizada"], sort=False):
        valor_original = _valor_unico_nao_nulo(grupo["VALOR"], titulo, data_efetiva)
        valor_classificado = sum(
            (valor_para_decimal(v) for v in grupo["VALOR_CLASS"]),
            start=Decimal("0.00"),
        )

        classificacoes = [
            Classificacao(
                codigo=str(linha["CLASSIFIC_COD"]),
                descricao=str(linha["CLASSIFIC_DESC"]),
                valor_class=valor_para_decimal(linha["VALOR_CLASS"]),
            )
            for _, linha in grupo.iterrows()
        ]

        pagamento = Pagamento(
            titulo=int(titulo),
            data_efetiva=data_efetiva,
            nome=_campo_invariante(grupo["NOME"], "NOME", titulo, data_efetiva),
            nf_num=_campo_preferencial(grupo["NF_NUM"]) or None,
            forma_pgto=_campo_invariante(grupo["FORMA_PGTO"], "FORMA_PGTO", titulo, data_efetiva),
            conta=_campo_invariante(grupo["CONTA"], "CONTA", titulo, data_efetiva),
            historico=_campo_preferencial(grupo["HISTORICO"]),
            valor_original=valor_original,
            valor_classificado=valor_classificado,
            classificacoes=classificacoes,
        )

        if pagamento.conta == CONTA_ALVO:
            candidatos.append(pagamento)
        else:
            fora_do_escopo.append(pagamento)

    return candidatos, fora_do_escopo


def listar_divergencias(pagamentos: list[Pagamento]) -> list[Pagamento]:
    """Filtra pagamentos cuja diferença VALOR x soma(VALOR_CLASS) passa a tolerância."""
    return [p for p in pagamentos if p.tem_divergencia_relevante]
