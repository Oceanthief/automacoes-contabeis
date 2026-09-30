"""Testes da Fase 1: leitura do extrato + padronização em 8 colunas.

Roda com o extrato do mês de exemplo (critérios de aceite da tarefa)
e com casos sintéticos pequenos para as regras de borda.
"""
import sys
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from leitura.extrato_itau import ExtratoItauError, filtrar_movimentacoes, ler_extrato
from fase1_padronizar_extrato import Fase1Error, executar, transformar, validar

ARQUIVO_AMOSTRA = RAIZ / "dados_exemplo" / "extrato_banco_2026_08.xlsx"


# ---------------------------------------------------------------------------
# Critérios de aceite explícitos da tarefa, com o extrato do mês de exemplo
# ---------------------------------------------------------------------------

def test_total_de_movimentacoes_exportadas():
    df = ler_extrato(ARQUIVO_AMOSTRA)
    mov = filtrar_movimentacoes(df)
    assert len(mov) == 58


def test_sequencia_vai_de_1_a_58():
    df = ler_extrato(ARQUIVO_AMOSTRA)
    mov = filtrar_movimentacoes(df)
    saida = transformar(mov)
    assert list(saida["sequencia"]) == list(range(1, 59))


def test_todas_as_linhas_tem_data_valor_descricao():
    df = ler_extrato(ARQUIVO_AMOSTRA)
    mov = filtrar_movimentacoes(df)
    saida = transformar(mov)
    assert not saida[["data", "valor", "descricao"]].isna().any().any()


def test_nenhum_lancamento_sem_valor_e_exportado():
    df = ler_extrato(ARQUIVO_AMOSTRA)
    mov = filtrar_movimentacoes(df)
    # Nenhuma linha de saldo (SALDO ANTERIOR / SALDO TOTAL DISPONÍVEL DIA) pode sobrar aqui.
    assert (~mov["Lançamento"].isin(["SALDO ANTERIOR", "SALDO TOTAL DISPONÍVEL DIA"])).sum() == len(mov) - 0
    assert mov["Valor (R$)"].notna().all()


def test_pagamentos_tem_credito_10001_e_recebimentos_debito_10001():
    df = ler_extrato(ARQUIVO_AMOSTRA)
    mov = filtrar_movimentacoes(df)
    saida = transformar(mov)

    pagamentos = saida[saida["valor"].notna()].merge(
        mov[["Valor (R$)"]].reset_index(drop=True), left_index=True, right_index=True
    )
    negativos = pagamentos[pagamentos["Valor (R$)"] < 0]
    positivos = pagamentos[pagamentos["Valor (R$)"] > 0]

    assert (negativos["credito"] == "10001").all()
    assert (negativos["debito"] == "").all()
    assert (positivos["debito"] == "10001").all()
    assert (positivos["credito"] == "").all()


def test_nunca_debito_e_credito_preenchidos_juntos():
    df = ler_extrato(ARQUIVO_AMOSTRA)
    mov = filtrar_movimentacoes(df)
    saida = transformar(mov)
    ambos = (saida["debito"] != "") & (saida["credito"] != "")
    assert not ambos.any()


def test_validar_nao_levanta_erro_para_arquivo_real():
    df = ler_extrato(ARQUIVO_AMOSTRA)
    mov = filtrar_movimentacoes(df)
    saida = transformar(mov)
    validar(saida)  # não deve lançar


# ---------------------------------------------------------------------------
# Teste específico pedido na tarefa: 03/08/2026, SISPAG FORNECEDORES, -1197,16
# ---------------------------------------------------------------------------

def test_caso_especifico_sispag_fornecedores():
    mov = pd.DataFrame(
        [{"Data": "03/08/2026", "Lançamento": "SISPAG FORNECEDORES", "Valor (R$)": -1197.16}]
    )
    saida = transformar(mov)
    linha = saida.iloc[0]

    assert linha["debito"] == ""
    assert linha["credito"] == "10001"
    assert linha["data"] == pd.Timestamp("2026-08-03")
    assert linha["valor"] == pytest.approx(1197.16)
    assert linha["descricao"] == "SISPAG FORNECEDORES"
    assert linha["sequencia"] == 1


# ---------------------------------------------------------------------------
# Regras de borda
# ---------------------------------------------------------------------------

def test_recebimento_gera_debito_sem_credito():
    mov = pd.DataFrame(
        [{"Data": "05/08/2026", "Lançamento": "PIX RECEBIDO TESTE", "Valor (R$)": 500.00}]
    )
    saida = transformar(mov)
    linha = saida.iloc[0]
    assert linha["debito"] == "10001"
    assert linha["credito"] == ""


def test_valor_sempre_exportado_positivo():
    mov = pd.DataFrame(
        [
            {"Data": "05/08/2026", "Lançamento": "PAGAMENTO", "Valor (R$)": -50.5},
            {"Data": "05/08/2026", "Lançamento": "RECEBIMENTO", "Valor (R$)": 50.5},
        ]
    )
    saida = transformar(mov)
    assert (saida["valor"] > 0).all()


def test_erro_se_arquivo_nao_existe():
    with pytest.raises(ExtratoItauError):
        ler_extrato("nao_existe.xlsx")


def test_erro_se_arquivo_saida_igual_a_origem():
    with pytest.raises(Fase1Error):
        executar(str(ARQUIVO_AMOSTRA), str(ARQUIVO_AMOSTRA))


def test_arquivo_original_nao_e_modificado(tmp_path):
    import hashlib

    hash_antes = hashlib.md5(ARQUIVO_AMOSTRA.read_bytes()).hexdigest()
    destino = tmp_path / "saida.xlsx"
    executar(str(ARQUIVO_AMOSTRA), str(destino))
    hash_depois = hashlib.md5(ARQUIVO_AMOSTRA.read_bytes()).hexdigest()

    assert hash_antes == hash_depois
    assert destino.exists()
