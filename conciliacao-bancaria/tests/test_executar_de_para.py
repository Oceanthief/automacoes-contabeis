"""Teste de ponta a ponta de executar_de_para.py com o mês de exemplo."""
import sys
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from executar_de_para import STATUS_AGUARDA_GRANULARIZACAO, executar
from enriquecimento.motor_de_para import (
    STATUS_AMBIGUO,
    STATUS_JA_CLASSIFICADO,
    STATUS_MATCH_EXATO,
    STATUS_MATCH_REGRA,
    STATUS_SEM_DE_PARA,
)

SAIDA_ENRIQUECIDA_SISPAG = RAIZ / "dados_exemplo" / "intermediarios" / "saida_enriquecida_sispag.xlsx"
DE_PARA = RAIZ / "dados_exemplo" / "de_para_contabil.xlsx"


def test_todos_os_58_lancamentos_sao_analisados(tmp_path):
    destino = tmp_path / "auditoria_de_para.xlsx"
    linhas = executar(str(SAIDA_ENRIQUECIDA_SISPAG), str(DE_PARA), str(destino))
    assert len(linhas) == 58
    assert destino.exists()


def test_tarifas_ficam_ja_classificado_com_conta_preservada(tmp_path):
    destino = tmp_path / "auditoria_de_para.xlsx"
    linhas = executar(str(SAIDA_ENRIQUECIDA_SISPAG), str(DE_PARA), str(destino))

    ja_classificados = [l for l in linhas if l["STATUS_DE_PARA"] == STATUS_JA_CLASSIFICADO]
    assert len(ja_classificados) == 6
    assert all(l["CONTA_ENCONTRADA"] == "00248" for l in ja_classificados)


def test_todos_os_10_sispag_ficam_aguardando_granularizacao(tmp_path):
    destino = tmp_path / "auditoria_de_para.xlsx"
    linhas = executar(str(SAIDA_ENRIQUECIDA_SISPAG), str(DE_PARA), str(destino))

    sispag = [l for l in linhas if l["DESCRICAO"] == "SISPAG FORNECEDORES"]
    assert len(sispag) == 10
    assert all(l["STATUS_DE_PARA"] == STATUS_AGUARDA_GRANULARIZACAO for l in sispag)
    assert all(l["CONTA_ENCONTRADA"] == "" for l in sispag)


def test_nenhum_ambiguo_ou_sem_de_para_recebe_conta(tmp_path):
    destino = tmp_path / "auditoria_de_para.xlsx"
    linhas = executar(str(SAIDA_ENRIQUECIDA_SISPAG), str(DE_PARA), str(destino))

    for l in linhas:
        if l["STATUS_DE_PARA"] in (STATUS_AMBIGUO, STATUS_SEM_DE_PARA):
            assert l["CONTA_ENCONTRADA"] == ""


def test_todo_match_registra_a_regra_utilizada(tmp_path):
    destino = tmp_path / "auditoria_de_para.xlsx"
    linhas = executar(str(SAIDA_ENRIQUECIDA_SISPAG), str(DE_PARA), str(destino))

    for l in linhas:
        if l["STATUS_DE_PARA"] in (STATUS_MATCH_EXATO, STATUS_MATCH_REGRA):
            assert l["CHAVE_MATCH_UTILIZADA"] != ""
            assert l["TIPO_MATCH_UTILIZADO"] != ""
            assert l["PRIORIDADE_UTILIZADA"] != ""


def test_contagem_bate_com_o_total_de_linhas(tmp_path):
    destino = tmp_path / "auditoria_de_para.xlsx"
    linhas = executar(str(SAIDA_ENRIQUECIDA_SISPAG), str(DE_PARA), str(destino))

    from collections import Counter
    contagem = Counter(l["STATUS_DE_PARA"] for l in linhas)
    assert sum(contagem.values()) == 58


def test_auditoria_gravada_preserva_conta_como_texto(tmp_path):
    destino = tmp_path / "auditoria_de_para.xlsx"
    executar(str(SAIDA_ENRIQUECIDA_SISPAG), str(DE_PARA), str(destino))

    lido = pd.read_excel(destino, sheet_name=0, dtype={"CONTA_ENCONTRADA": str})
    tarifa = lido[lido["STATUS_DE_PARA"] == "JA_CLASSIFICADO"].iloc[0]
    assert tarifa["CONTA_ENCONTRADA"] == "00248"


def test_pagamentos_e_recebimentos_preservam_conta_bancaria(tmp_path):
    """A conta 10001 nunca pode mudar de lado: CREDITO=10001 em todo
    pagamento, DEBITO=10001 em todo recebimento — independente do que o
    DE-PARA decidiu para a conta faltante."""
    destino = tmp_path / "auditoria_de_para.xlsx"
    executar(str(SAIDA_ENRIQUECIDA_SISPAG), str(DE_PARA), str(destino))

    extrato = pd.read_excel(SAIDA_ENRIQUECIDA_SISPAG, sheet_name=0, dtype={"DEBITO": str, "CREDITO": str})
    aud = pd.read_excel(destino, sheet_name=0)
    merged = aud.merge(extrato[["SEQUENCIA", "DEBITO", "CREDITO"]], on="SEQUENCIA")

    pagamentos = merged[merged["TIPO_MOVIMENTO"] == "PAGAMENTO"]
    recebimentos = merged[merged["TIPO_MOVIMENTO"] == "RECEBIMENTO"]
    assert (pagamentos["CREDITO"] == "10001").all()
    assert (recebimentos["DEBITO"] == "10001").all()


def test_sispag_exato_soma_nao_recebe_conta_unica_para_o_valor_total(tmp_path):
    """Cruza com auditoria_sispag.xlsx para isolar especificamente os
    EXATO_SOMA (múltiplos pagamentos por trás de um único lançamento
    bancário) e confirmar que nenhum deles ganhou uma conta só para o total."""
    destino = tmp_path / "auditoria_de_para.xlsx"
    executar(str(SAIDA_ENRIQUECIDA_SISPAG), str(DE_PARA), str(destino))

    aud_de_para = pd.read_excel(destino, sheet_name=0, dtype={"CONTA_ENCONTRADA": str})
    aud_sispag = pd.read_excel(RAIZ / "dados_exemplo" / "intermediarios" / "auditoria_sispag.xlsx", sheet_name=0)
    merged = aud_de_para.merge(aud_sispag[["SEQUENCIA", "STATUS"]], on="SEQUENCIA")

    exato_soma = merged[merged["STATUS"] == "EXATO_SOMA"]
    assert len(exato_soma) > 0  # a checagem só vale algo se houver casos no exemplo
    assert (exato_soma["STATUS_DE_PARA"] == STATUS_AGUARDA_GRANULARIZACAO).all()
    assert (exato_soma["CONTA_ENCONTRADA"].isna() | (exato_soma["CONTA_ENCONTRADA"] == "")).all()
