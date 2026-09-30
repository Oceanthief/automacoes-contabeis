"""Testes de enriquecimento/classificacao_extrato.py."""
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from enriquecimento.classificacao_extrato import (
    CONTA_TARIFA,
    ClassificacaoError,
    classificar_linha,
    eh_tarifa_bancaria,
)


# ---------------------------------------------------------------------------
# Regra 1: prefixo de pagamento/recebimento
# ---------------------------------------------------------------------------

def test_pagamento_recebe_prefixo_pgto():
    r = classificar_linha("BOLETO FORNECEDOR X", debito=float("nan"), credito=10001.0)
    assert r.nova_descricao == "VLR REF. PGTO BOLETO FORNECEDOR X"
    assert r.eh_pagamento is True


def test_recebimento_recebe_prefixo_recebimento():
    r = classificar_linha("PIX RECEBIDO FULANO", debito=10001.0, credito=float("nan"))
    assert r.nova_descricao == "VLR REF. RECEBIMENTO PIX RECEBIDO FULANO"
    assert r.eh_pagamento is False


def test_descricao_original_preservada_na_nova_descricao():
    r = classificar_linha("  Texto Com Espaco  ", debito=float("nan"), credito=10001.0)
    # a concatenação usa a descrição como veio (só passada por str().strip()
    # no nível do próprio texto original, sem maiúsculas/acentos removidos —
    # isso é só para a NOVA DESCRICAO, nunca para DESCRICAO em si)
    assert r.nova_descricao == "VLR REF. PGTO Texto Com Espaco"


# ---------------------------------------------------------------------------
# Exceção: SISPAG FORNECEDORES
# ---------------------------------------------------------------------------

def test_sispag_fornecedores_fica_com_nova_descricao_vazia():
    r = classificar_linha("SISPAG FORNECEDORES", debito=float("nan"), credito=10001.0)
    assert r.nova_descricao == ""
    assert r.eh_tarifa is False


# ---------------------------------------------------------------------------
# Regra 2: tarifas bancárias
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "descricao",
    [
        "TAR SISPAG CONCESSION",
        "TAR CTA EMP MENSAL 07/26",
        "TAR DEPOSITO DIN AGENCIA",
        "TAR PIX PGTO TRANSF",
        "TAR BLOQUETO ITAU",
        "TAR SISPAG TIT OUTRO BCO",
        "TAR COBRANCA MENSAL",
    ],
)
def test_exemplos_de_tarifa_sao_reconhecidos(descricao):
    assert eh_tarifa_bancaria(descricao, debito=float("nan"), credito=10001.0) is True
    r = classificar_linha(descricao, debito=float("nan"), credito=10001.0)
    assert r.debito == CONTA_TARIFA
    assert r.credito == "10001"


def test_tarifa_com_prefixo_tarifa_por_extenso_tambem_e_reconhecida():
    assert eh_tarifa_bancaria("TARIFA MANUTENCAO CONTA", debito=float("nan"), credito=10001.0) is True


def test_bo_est_lanc_tarifa_nao_e_classificado_como_tarifa_de_saida():
    """Caso do enunciado: é um RECEBIMENTO, não pode virar despesa de tarifa."""
    assert eh_tarifa_bancaria("BO  EST LANC TARIFA", debito=10001.0, credito=float("nan")) is False
    r = classificar_linha("BO  EST LANC TARIFA", debito=10001.0, credito=float("nan"))
    assert r.eh_tarifa is False
    assert r.debito == "10001"
    assert r.nova_descricao == "VLR REF. RECEBIMENTO BO  EST LANC TARIFA"


def test_palavra_parecida_com_tar_nao_e_tarifa_por_substring():
    """'TARDES' ou algo que comece com TAR mas não seja a palavra TAR/TARIFA
    isolada não deveria disparar a regra (prefixo/palavra, não substring)."""
    assert eh_tarifa_bancaria("TARDES PAGAMENTOS DIVERSOS", debito=float("nan"), credito=10001.0) is False


def test_recebimento_com_tar_no_inicio_nunca_vira_tarifa():
    assert eh_tarifa_bancaria("TAR ALGO", debito=10001.0, credito=float("nan")) is False


def test_normalizacao_ignora_acento_caixa_e_espacos_duplos():
    assert eh_tarifa_bancaria("  tar  depósito  ", debito=float("nan"), credito=10001.0) is True


# ---------------------------------------------------------------------------
# Robustez de tipos de DEBITO/CREDITO (float com .0, NaN, string)
# ---------------------------------------------------------------------------

def test_aceita_conta_como_string_ou_float():
    r_float = classificar_linha("PAGAMENTO TESTE", debito=float("nan"), credito=10001.0)
    r_str = classificar_linha("PAGAMENTO TESTE", debito="", credito="10001")
    assert r_float.eh_pagamento == r_str.eh_pagamento == True


def test_linha_sem_debito_nem_credito_levanta_erro():
    with pytest.raises(ClassificacaoError):
        classificar_linha("LINHA QUEBRADA", debito=float("nan"), credito=float("nan"))


def test_linha_com_debito_e_credito_simultaneos_levanta_erro():
    with pytest.raises(ClassificacaoError):
        classificar_linha("LINHA QUEBRADA", debito=10001.0, credito=10001.0)
