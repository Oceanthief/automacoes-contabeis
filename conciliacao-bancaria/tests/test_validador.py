"""Testes de validacao/validador.py."""
import sys
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from validacao.validador import (
    STATUS_APROVADO,
    STATUS_BLOQUEADO,
    STATUS_GERAL_APROVADO,
    STATUS_GERAL_BLOQUEADO,
    calcular_totais,
    status_geral,
    validar_saida,
)

COLUNAS_BASE = [
    "SEQUENCIA_CONTABIL", "SEQUENCIA_BANCO", "DEBITO", "CREDITO", "DATA", "VALOR",
    "DESCRICAO_ORIGINAL", "NOVA_DESCRICAO", "TIPO_MOVIMENTO", "TIPO_ORIGEM",
    "STATUS_CONCILIACAO", "STATUS_DE_PARA", "TITULO", "FORNECEDOR", "NF_NUM",
    "HISTORICO", "FORMA_PGTO", "CONTA_ORIGEM",
]


def _linha(**overrides):
    base = {
        "SEQUENCIA_CONTABIL": 1, "SEQUENCIA_BANCO": 1, "DEBITO": "83860", "CREDITO": "10001",
        "DATA": pd.Timestamp("2026-08-03"), "VALOR": 100.0,
        "DESCRICAO_ORIGINAL": "BOLETO FORNECEDOR", "NOVA_DESCRICAO": "VLR REF. PGTO FORNECEDOR X",
        "TIPO_MOVIMENTO": "PAGAMENTO", "TIPO_ORIGEM": "NORMAL", "STATUS_CONCILIACAO": "",
        "STATUS_DE_PARA": "MATCH_EXATO", "TITULO": "", "FORNECEDOR": "", "NF_NUM": "",
        "HISTORICO": "", "FORMA_PGTO": "", "CONTA_ORIGEM": "",
    }
    base.update(overrides)
    return base


def _df(*linhas):
    return pd.DataFrame(list(linhas))


AUDITORIA_GRAN_VAZIA = pd.DataFrame(columns=["SEQUENCIA_BANCO", "STATUS_GRANULARIZACAO", "STATUS_CONCILIACAO"])


# ---------------------------------------------------------------------------
# Caso feliz
# ---------------------------------------------------------------------------

def test_linha_valida_e_aprovada():
    df = _df(_linha())
    resultado, contagem = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert resultado.iloc[0]["STATUS_VALIDACAO"] == STATUS_APROVADO
    assert resultado.iloc[0]["MOTIVO_BLOQUEIO"] == ""
    assert status_geral(resultado) == STATUS_GERAL_APROVADO


# ---------------------------------------------------------------------------
# 1) DATA
# ---------------------------------------------------------------------------

def test_data_vazia_bloqueia():
    df = _df(_linha(DATA=None))
    resultado, _ = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert resultado.iloc[0]["STATUS_VALIDACAO"] == STATUS_BLOQUEADO
    assert "DATA" in resultado.iloc[0]["MOTIVO_BLOQUEIO"]


def test_data_fora_da_competencia_bloqueia():
    # 9 linhas em agosto, 1 linha isolada em outro mês
    linhas = [_linha(SEQUENCIA_CONTABIL=i, SEQUENCIA_BANCO=i) for i in range(1, 10)]
    linhas.append(_linha(SEQUENCIA_CONTABIL=10, SEQUENCIA_BANCO=10, DATA=pd.Timestamp("2026-01-15")))
    df = _df(*linhas)
    resultado, _ = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    fora = resultado[resultado["SEQUENCIA_CONTABIL"] == 10].iloc[0]
    assert fora["STATUS_VALIDACAO"] == STATUS_BLOQUEADO
    assert "competência" in fora["MOTIVO_BLOQUEIO"]


# ---------------------------------------------------------------------------
# 2) VALOR
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("valor", [0.0, -10.0, None])
def test_valor_invalido_bloqueia(valor):
    df = _df(_linha(VALOR=valor))
    resultado, _ = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert resultado.iloc[0]["STATUS_VALIDACAO"] == STATUS_BLOQUEADO


# ---------------------------------------------------------------------------
# 3) DEBITO/CREDITO preenchidos
# ---------------------------------------------------------------------------

def test_debito_vazio_bloqueia():
    df = _df(_linha(DEBITO="", STATUS_DE_PARA="SEM_DE_PARA"))
    resultado, _ = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert resultado.iloc[0]["STATUS_VALIDACAO"] == STATUS_BLOQUEADO
    assert "DEBITO vazio" in resultado.iloc[0]["MOTIVO_BLOQUEIO"]


# ---------------------------------------------------------------------------
# 4) Conta bancária
# ---------------------------------------------------------------------------

def test_pagamento_com_credito_diferente_de_10001_bloqueia():
    df = _df(_linha(CREDITO="99999"))
    resultado, _ = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert resultado.iloc[0]["STATUS_VALIDACAO"] == STATUS_BLOQUEADO
    assert "CREDITO diferente" in resultado.iloc[0]["MOTIVO_BLOQUEIO"]


def test_recebimento_com_debito_diferente_de_10001_bloqueia():
    df = _df(_linha(TIPO_MOVIMENTO="RECEBIMENTO", DEBITO="99999", CREDITO="05620"))
    resultado, _ = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert resultado.iloc[0]["STATUS_VALIDACAO"] == STATUS_BLOQUEADO


# ---------------------------------------------------------------------------
# 5) DEBITO == CREDITO
# ---------------------------------------------------------------------------

def test_debito_igual_credito_bloqueia():
    df = _df(_linha(DEBITO="10001", CREDITO="10001"))
    resultado, _ = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert "ERRO_CONTAS_IGUAIS" in resultado.iloc[0]["MOTIVO_BLOQUEIO"]


# ---------------------------------------------------------------------------
# 6) DE-PARA
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("status", ["SEM_DE_PARA", "AMBIGUO"])
def test_status_de_para_bloqueante(status):
    df = _df(_linha(STATUS_DE_PARA=status, DEBITO=""))
    resultado, _ = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert resultado.iloc[0]["STATUS_VALIDACAO"] == STATUS_BLOQUEADO


def test_ajuste_manual_nao_bloqueia_por_de_para():
    df = _df(_linha(STATUS_DE_PARA="AJUSTE_MANUAL"))
    resultado, _ = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert resultado.iloc[0]["STATUS_VALIDACAO"] == STATUS_APROVADO


# ---------------------------------------------------------------------------
# 7) Regra fixa de tarifa
# ---------------------------------------------------------------------------

def test_tarifa_com_conta_correta_aprovada():
    df = _df(_linha(DESCRICAO_ORIGINAL="TAR BLOQUETO ITAU", DEBITO="00248", CREDITO="10001", STATUS_DE_PARA="JA_CLASSIFICADO"))
    resultado, _ = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert resultado.iloc[0]["STATUS_VALIDACAO"] == STATUS_APROVADO


def test_tarifa_com_conta_errada_bloqueia():
    df = _df(_linha(DESCRICAO_ORIGINAL="TAR BLOQUETO ITAU", DEBITO="99999", CREDITO="10001", STATUS_DE_PARA="JA_CLASSIFICADO"))
    resultado, _ = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert resultado.iloc[0]["STATUS_VALIDACAO"] == STATUS_BLOQUEADO
    assert "Tarifa" in resultado.iloc[0]["MOTIVO_BLOQUEIO"]


# ---------------------------------------------------------------------------
# 8) SISPAG
# ---------------------------------------------------------------------------

def test_sispag_com_erro_granularizacao_bloqueia_linhas_daquele_movimento():
    df = _df(
        _linha(SEQUENCIA_CONTABIL=1, SEQUENCIA_BANCO=50, DESCRICAO_ORIGINAL="SISPAG FORNECEDORES", TITULO="123"),
        _linha(SEQUENCIA_CONTABIL=2, SEQUENCIA_BANCO=50, DESCRICAO_ORIGINAL="SISPAG FORNECEDORES", TITULO="124"),
    )
    auditoria_gran = pd.DataFrame(
        [{"SEQUENCIA_BANCO": 50, "STATUS_GRANULARIZACAO": "ERRO_GRANULARIZACAO", "STATUS_CONCILIACAO": "EXATO_SOMA"}]
    )
    resultado, contagem = validar_saida(df, auditoria_gran)
    assert (resultado["STATUS_VALIDACAO"] == STATUS_BLOQUEADO).all()
    assert contagem["SISPAG"] == 2


def test_sispag_ok_nao_bloqueia():
    df = _df(_linha(SEQUENCIA_BANCO=50, DESCRICAO_ORIGINAL="SISPAG FORNECEDORES", TITULO="123"))
    auditoria_gran = pd.DataFrame(
        [{"SEQUENCIA_BANCO": 50, "STATUS_GRANULARIZACAO": "OK", "STATUS_CONCILIACAO": "EXATO_1X1"}]
    )
    resultado, _ = validar_saida(df, auditoria_gran)
    assert resultado.iloc[0]["STATUS_VALIDACAO"] == STATUS_APROVADO


# ---------------------------------------------------------------------------
# 9) Duplicidade
# ---------------------------------------------------------------------------

def test_titulo_duplicado_bloqueia_ambas_as_linhas():
    df = _df(
        _linha(SEQUENCIA_CONTABIL=1, SEQUENCIA_BANCO=10, TITULO="999"),
        _linha(SEQUENCIA_CONTABIL=2, SEQUENCIA_BANCO=11, TITULO="999"),
    )
    resultado, contagem = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert (resultado["STATUS_VALIDACAO"] == STATUS_BLOQUEADO).all()
    assert contagem["DUPLICIDADE"] == 2


def test_fornecedor_repetido_com_titulos_diferentes_nao_e_duplicidade():
    df = _df(
        _linha(SEQUENCIA_CONTABIL=1, SEQUENCIA_BANCO=10, TITULO="111", FORNECEDOR="Fornecedor A"),
        _linha(SEQUENCIA_CONTABIL=2, SEQUENCIA_BANCO=11, TITULO="222", FORNECEDOR="Fornecedor A"),
    )
    resultado, contagem = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert (resultado["STATUS_VALIDACAO"] == STATUS_APROVADO).all()
    assert contagem["DUPLICIDADE"] == 0


def test_sequencia_contabil_duplicada_levanta_erro():
    df = _df(_linha(SEQUENCIA_CONTABIL=1), _linha(SEQUENCIA_CONTABIL=1))
    with pytest.raises(ValueError):
        validar_saida(df, AUDITORIA_GRAN_VAZIA)


# ---------------------------------------------------------------------------
# 11) Descrição
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("descricao", ["", "   ", None])
def test_nova_descricao_vazia_bloqueia(descricao):
    df = _df(_linha(NOVA_DESCRICAO=descricao))
    resultado, _ = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert resultado.iloc[0]["STATUS_VALIDACAO"] == STATUS_BLOQUEADO


# ---------------------------------------------------------------------------
# 12) Totalização — granularização não altera valor econômico
# ---------------------------------------------------------------------------

def test_totais_pagamento_recebimento():
    df = _df(
        _linha(SEQUENCIA_CONTABIL=1, SEQUENCIA_BANCO=1, TIPO_MOVIMENTO="PAGAMENTO", VALOR=100.0),
        _linha(SEQUENCIA_CONTABIL=2, SEQUENCIA_BANCO=2, TIPO_MOVIMENTO="RECEBIMENTO", DEBITO="10001", CREDITO="05620", VALOR=50.0),
    )
    totais = calcular_totais(df)
    assert totais.total_pagamentos == 100
    assert totais.total_recebimentos == 50
    assert totais.total_debitos == totais.total_creditos == 150


def test_totais_sispag_granularizado_nao_duplica_valor_bancario():
    df = _df(
        _linha(SEQUENCIA_CONTABIL=1, SEQUENCIA_BANCO=10, DESCRICAO_ORIGINAL="SISPAG FORNECEDORES", VALOR=60.0),
        _linha(SEQUENCIA_CONTABIL=2, SEQUENCIA_BANCO=10, DESCRICAO_ORIGINAL="SISPAG FORNECEDORES", VALOR=40.0),
    )
    auditoria_gran = pd.DataFrame([{"SEQUENCIA_BANCO": 10, "VALOR_BANCO": 100.0}])
    totais = calcular_totais(df, auditoria_gran)
    assert totais.total_sispag_granularizado == 100  # soma das 2 linhas
    assert totais.total_bancario_sispag == 100  # 1 movimento bancário só, não 2


# ---------------------------------------------------------------------------
# STATUS_GERAL
# ---------------------------------------------------------------------------

def test_status_geral_bloqueado_se_alguma_linha_bloqueada():
    df = _df(_linha(SEQUENCIA_CONTABIL=1), _linha(SEQUENCIA_CONTABIL=2, VALOR=0.0))
    resultado, _ = validar_saida(df, AUDITORIA_GRAN_VAZIA)
    assert status_geral(resultado) == STATUS_GERAL_BLOQUEADO
