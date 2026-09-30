"""Testes de validacao/aplicar_correcoes.py."""
import sys
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from validacao.aplicar_correcoes import AplicarCorrecoesError, STATUS_AJUSTE_MANUAL, aplicar_correcoes


def _linha_saida(seq, tipo_movimento="PAGAMENTO", debito="", credito="10001", status_de_para="SEM_DE_PARA"):
    return {
        "SEQUENCIA_CONTABIL": seq, "TIPO_MOVIMENTO": tipo_movimento,
        "DEBITO": debito, "CREDITO": credito, "STATUS_DE_PARA": status_de_para,
    }


def test_corrige_pagamento_sem_de_para():
    saida = pd.DataFrame([_linha_saida(1)])
    correcoes = pd.DataFrame([{"SEQUENCIA_CONTABIL": 1, "DEBITO_MANUAL": "83860", "CREDITO_MANUAL": "10001", "OBSERVACAO": ""}])

    resultado = aplicar_correcoes(saida, correcoes)

    linha = resultado.iloc[0]
    assert linha["DEBITO"] == "83860"
    assert linha["STATUS_DE_PARA"] == STATUS_AJUSTE_MANUAL
    assert linha["STATUS_DE_PARA_ORIGINAL"] == "SEM_DE_PARA"


def test_corrige_ambiguo_tambem_e_permitido():
    saida = pd.DataFrame([_linha_saida(1, status_de_para="AMBIGUO")])
    correcoes = pd.DataFrame([{"SEQUENCIA_CONTABIL": 1, "DEBITO_MANUAL": "12345", "CREDITO_MANUAL": "10001", "OBSERVACAO": ""}])

    resultado = aplicar_correcoes(saida, correcoes)
    assert resultado.iloc[0]["DEBITO"] == "12345"


def test_bloqueia_correcao_sobre_linha_ja_classificada():
    saida = pd.DataFrame([_linha_saida(1, status_de_para="MATCH_EXATO", debito="99999")])
    correcoes = pd.DataFrame([{"SEQUENCIA_CONTABIL": 1, "DEBITO_MANUAL": "12345", "CREDITO_MANUAL": "10001", "OBSERVACAO": ""}])

    with pytest.raises(AplicarCorrecoesError):
        aplicar_correcoes(saida, correcoes)


def test_bloqueia_sequencia_inexistente():
    saida = pd.DataFrame([_linha_saida(1)])
    correcoes = pd.DataFrame([{"SEQUENCIA_CONTABIL": 999, "DEBITO_MANUAL": "12345", "CREDITO_MANUAL": "10001", "OBSERVACAO": ""}])

    with pytest.raises(AplicarCorrecoesError):
        aplicar_correcoes(saida, correcoes)


def test_bloqueia_credito_manual_divergente():
    saida = pd.DataFrame([_linha_saida(1, credito="10001")])
    correcoes = pd.DataFrame([{"SEQUENCIA_CONTABIL": 1, "DEBITO_MANUAL": "12345", "CREDITO_MANUAL": "99999", "OBSERVACAO": ""}])

    with pytest.raises(AplicarCorrecoesError):
        aplicar_correcoes(saida, correcoes)


def test_recebimento_usa_credito_manual():
    saida = pd.DataFrame([_linha_saida(1, tipo_movimento="RECEBIMENTO", debito="10001", credito="")])
    correcoes = pd.DataFrame([{"SEQUENCIA_CONTABIL": 1, "DEBITO_MANUAL": "", "CREDITO_MANUAL": "05620", "OBSERVACAO": ""}])

    resultado = aplicar_correcoes(saida, correcoes)
    assert resultado.iloc[0]["CREDITO"] == "05620"


def test_nao_altera_dataframe_original():
    saida = pd.DataFrame([_linha_saida(1)])
    correcoes = pd.DataFrame([{"SEQUENCIA_CONTABIL": 1, "DEBITO_MANUAL": "83860", "CREDITO_MANUAL": "10001", "OBSERVACAO": ""}])

    aplicar_correcoes(saida, correcoes)
    assert saida.iloc[0]["DEBITO"] == ""  # original intocado
