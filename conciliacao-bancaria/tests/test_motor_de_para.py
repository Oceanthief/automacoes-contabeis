"""Testes do motor de DE-PARA (enriquecimento/motor_de_para.py)."""
import sys
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from enriquecimento.motor_de_para import (
    STATUS_AMBIGUO,
    STATUS_MATCH_EXATO,
    STATUS_MATCH_REGRA,
    STATUS_SEM_DE_PARA,
    MotorDeParaError,
    carregar_regras,
    classificar,
)
from normalizacao import normalizar_texto_classificacao


def _regra(chave, tipo_movimento="PAGAMENTO", tipo_match="EXATO", conta="12345", prioridade=1, ativo="SIM",
           historico="", observacao=""):
    return {
        "CHAVE_MATCH": chave, "TIPO_MOVIMENTO": tipo_movimento, "TIPO_MATCH": tipo_match,
        "CONTA_CONTABIL": conta, "PRIORIDADE": prioridade, "ATIVO": ativo,
        "HISTORICO_PADRAO": historico, "OBSERVACAO": observacao,
    }


def _regras(*linhas):
    return carregar_regras(pd.DataFrame(list(linhas)))


# ---------------------------------------------------------------------------
# Tipos de match básicos
# ---------------------------------------------------------------------------

def test_match_exato():
    regras = _regras(_regra("ESCRITORIO CONTABIL MODELO", conta="02997"))
    r = classificar(regras, "PAGAMENTO", normalizar_texto_classificacao("ESCRITORIO CONTABIL MODELO"))
    assert r.status == STATUS_MATCH_EXATO
    assert r.conta_encontrada == "02997"


def test_match_comeca_com():
    regras = _regras(_regra("SILVA & SOUZA", tipo_match="COMECA_COM", conta="55555"))
    texto = normalizar_texto_classificacao("SILVA & SOUZA COMERCIO DE TINTAS LTDA")
    r = classificar(regras, "PAGAMENTO", texto)
    assert r.status == STATUS_MATCH_REGRA
    assert r.conta_encontrada == "55555"


def test_match_contem():
    regras = _regras(_regra("SEGUROS", tipo_match="CONTEM", conta="77777"))
    texto = normalizar_texto_classificacao("HORIZONTE SEGUROS S/A")
    r = classificar(regras, "PAGAMENTO", texto)
    assert r.status == STATUS_MATCH_REGRA
    assert r.conta_encontrada == "77777"


def test_sem_de_para_quando_nada_bate():
    regras = _regras(_regra("XPTO"))
    r = classificar(regras, "PAGAMENTO", normalizar_texto_classificacao("FORNECEDOR DESCONHECIDO"))
    assert r.status == STATUS_SEM_DE_PARA


def test_texto_match_vazio_e_sem_de_para():
    regras = _regras(_regra("XPTO"))
    r = classificar(regras, "PAGAMENTO", "")
    assert r.status == STATUS_SEM_DE_PARA


# ---------------------------------------------------------------------------
# Sem fuzzy matching: variações não batem
# ---------------------------------------------------------------------------

def test_nao_faz_fuzzy_matching():
    regras = _regras(_regra("COMPANHIA DE SANEAMENTO DO VALE AGUAS SA"))
    texto = normalizar_texto_classificacao("AGUAS SA SANEAMENTO")  # ordem diferente, não é EXATO nem CONTEM
    r = classificar(regras, "PAGAMENTO", texto)
    assert r.status == STATUS_SEM_DE_PARA


# ---------------------------------------------------------------------------
# TIPO_MOVIMENTO: PAGAMENTO / RECEBIMENTO / AMBOS
# ---------------------------------------------------------------------------

def test_regra_de_pagamento_nao_serve_para_recebimento():
    regras = _regras(_regra("FORNECEDOR X", tipo_movimento="PAGAMENTO"))
    texto = normalizar_texto_classificacao("FORNECEDOR X")
    r = classificar(regras, "RECEBIMENTO", texto)
    assert r.status == STATUS_SEM_DE_PARA


def test_regra_ambos_serve_para_pagamento_e_recebimento():
    regras = _regras(_regra("CLIENTE Y", tipo_movimento="AMBOS", conta="999"))
    texto = normalizar_texto_classificacao("CLIENTE Y")
    assert classificar(regras, "PAGAMENTO", texto).status == STATUS_MATCH_EXATO
    assert classificar(regras, "RECEBIMENTO", texto).status == STATUS_MATCH_EXATO


# ---------------------------------------------------------------------------
# ATIVO = NAO nunca classifica
# ---------------------------------------------------------------------------

def test_regra_inativa_nunca_classifica():
    regras = _regras(_regra("FORNECEDOR INATIVO", ativo="NAO"))
    texto = normalizar_texto_classificacao("FORNECEDOR INATIVO")
    r = classificar(regras, "PAGAMENTO", texto)
    assert r.status == STATUS_SEM_DE_PARA


# ---------------------------------------------------------------------------
# Ordem: prioridade numérica primeiro, especificidade como desempate
# ---------------------------------------------------------------------------

def test_prioridade_numerica_vence_mesmo_com_tipo_match_menos_especifico():
    # CONTEM com prioridade 1 deveria vencer EXATO com prioridade 2
    regras = _regras(
        _regra("PARTE", tipo_match="CONTEM", conta="111", prioridade=1),
        _regra("FORNECEDOR PARTE DO NOME", tipo_match="EXATO", conta="222", prioridade=2),
    )
    texto = normalizar_texto_classificacao("FORNECEDOR PARTE DO NOME")
    r = classificar(regras, "PAGAMENTO", texto)
    assert r.conta_encontrada == "111"


def test_especificidade_desempata_prioridade_igual():
    regras = _regras(
        _regra("FORNECEDOR", tipo_match="CONTEM", conta="111", prioridade=1),
        _regra("FORNECEDOR X", tipo_match="EXATO", conta="222", prioridade=1),
    )
    texto = normalizar_texto_classificacao("FORNECEDOR X")
    r = classificar(regras, "PAGAMENTO", texto)
    assert r.status == STATUS_MATCH_EXATO
    assert r.conta_encontrada == "222"


# ---------------------------------------------------------------------------
# Ambiguidade
# ---------------------------------------------------------------------------

def test_duas_regras_empatadas_mesma_conta_nao_e_ambiguo():
    regras = _regras(
        _regra("FORNECEDOR Z", tipo_match="EXATO", conta="333", prioridade=1),
        _regra("FORNECEDOR Z ", tipo_match="EXATO", conta="333", prioridade=1),  # normaliza igual
    )
    texto = normalizar_texto_classificacao("FORNECEDOR Z")
    r = classificar(regras, "PAGAMENTO", texto)
    assert r.status == STATUS_MATCH_EXATO
    assert r.conta_encontrada == "333"


def test_duas_regras_empatadas_contas_diferentes_e_ambiguo():
    regras = _regras(
        _regra("FORNECEDOR W", tipo_match="EXATO", conta="111", prioridade=1),
        _regra("FORNECEDOR W", tipo_match="EXATO", conta="222", prioridade=1),
    )
    texto = normalizar_texto_classificacao("FORNECEDOR W")
    r = classificar(regras, "PAGAMENTO", texto)
    assert r.status == STATUS_AMBIGUO
    assert r.conta_encontrada is None


# ---------------------------------------------------------------------------
# Validação de regras malformadas
# ---------------------------------------------------------------------------

def test_tipo_movimento_invalido_levanta_erro():
    with pytest.raises(MotorDeParaError):
        _regras(_regra("X", tipo_movimento="SAIDA"))


def test_tipo_match_invalido_levanta_erro():
    with pytest.raises(MotorDeParaError):
        _regras(_regra("X", tipo_match="PARECIDO"))


def test_ativo_invalido_levanta_erro():
    with pytest.raises(MotorDeParaError):
        _regras(_regra("X", ativo="TALVEZ"))


def test_prioridade_ausente_levanta_erro():
    with pytest.raises(MotorDeParaError):
        _regras(_regra("X", prioridade=None))


def test_conta_placeholder_em_regra_ativa_levanta_erro():
    with pytest.raises(MotorDeParaError):
        _regras(_regra("X", conta="<preencher>", ativo="SIM"))


def test_conta_placeholder_em_regra_inativa_nao_levanta_erro():
    regras = _regras(_regra("X", conta="<preencher>", ativo="NAO"))
    assert len(regras) == 1
    assert regras[0].ativo is False
