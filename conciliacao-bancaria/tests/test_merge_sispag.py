"""Testes do merge controlado entre saida_enriquecida e auditoria_sispag."""
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from enriquecimento.merge_sispag import MergeSispagError, _dedupe_nomes, processar


def _linha_saida(sequencia, data, valor, descricao="SISPAG FORNECEDORES"):
    return {"SEQUENCIA": sequencia, "DATA": pd.Timestamp(data), "VALOR": valor, "DESCRICAO": descricao}


def _linha_auditoria(sequencia, data, valor_sispag, status, nomes=""):
    return {
        "SEQUENCIA": sequencia, "DATA": pd.Timestamp(data), "VALOR_SISPAG": valor_sispag,
        "STATUS": status, "NOMES": nomes,
    }


# ---------------------------------------------------------------------------
# Dedupe de nomes (regra central desta fase)
# ---------------------------------------------------------------------------

def test_dedupe_remove_duplicado_exato():
    entrada = "COMPANHIA DE SANEAMENTO DO VALE AGUAS SA; COMPANHIA DE SANEAMENTO DO VALE AGUAS SA"
    assert _dedupe_nomes(entrada) == "COMPANHIA DE SANEAMENTO DO VALE AGUAS SA"


def test_dedupe_preserva_multiplos_fornecedores_distintos():
    entrada = "Fornecedor A; Fornecedor B; Fornecedor A"
    assert _dedupe_nomes(entrada) == "Fornecedor A; Fornecedor B"


def test_dedupe_ignora_espacos_excedentes():
    entrada = "  Fornecedor X  ;   Fornecedor X  "
    assert _dedupe_nomes(entrada) == "Fornecedor X"


def test_dedupe_compara_normalizado_mas_preserva_grafia_original():
    entrada = "fornecedor y; FORNECEDOR Y; Fornecedor Y"
    # a primeira grafia encontrada é "fornecedor y" (minúsculo)
    assert _dedupe_nomes(entrada) == "fornecedor y"


def test_dedupe_preserva_ordem_original():
    entrada = "Zeta; Alfa; Beta"
    assert _dedupe_nomes(entrada) == "Zeta; Alfa; Beta"


# ---------------------------------------------------------------------------
# Merge: casos felizes
# ---------------------------------------------------------------------------

def test_exato_1x1_preenche_com_um_fornecedor():
    saida = pd.DataFrame([_linha_saida(4, "2026-08-03", 333.00)])
    auditoria = pd.DataFrame([_linha_auditoria(4, "2026-08-03", 333.00, "EXATO_1X1", "Fornecedor Único")])

    resultados = processar(saida, auditoria)
    assert len(resultados) == 1
    assert resultados[0].preenchido is True
    assert resultados[0].nova_descricao == "Fornecedor Único"


def test_exato_soma_preenche_com_multiplos_fornecedores_sem_duplicata():
    saida = pd.DataFrame([_linha_saida(3, "2026-08-03", 1197.16)])
    nomes = "AGUAS SA; AGUAS SA"
    auditoria = pd.DataFrame([_linha_auditoria(3, "2026-08-03", 1197.16, "EXATO_SOMA", nomes)])

    resultados = processar(saida, auditoria)
    assert resultados[0].preenchido is True
    assert resultados[0].nova_descricao == "AGUAS SA"


def test_linhas_nao_sispag_sao_ignoradas():
    saida = pd.DataFrame([_linha_saida(1, "2026-08-03", 500.0, descricao="PIX RECEBIDO")])
    auditoria = pd.DataFrame([_linha_auditoria(1, "2026-08-03", 500.0, "EXATO_1X1", "X")])

    resultados = processar(saida, auditoria)
    assert resultados == []


# ---------------------------------------------------------------------------
# Merge: casos de erro/divergência
# ---------------------------------------------------------------------------

def test_sequencia_sem_correspondencia_nao_preenche():
    saida = pd.DataFrame([_linha_saida(999, "2026-08-03", 100.0)])
    auditoria = pd.DataFrame([_linha_auditoria(1, "2026-08-03", 100.0, "EXATO_1X1", "X")])

    resultados = processar(saida, auditoria)
    assert resultados[0].preenchido is False
    assert "sem correspondência" in resultados[0].motivo


def test_divergencia_de_data_nao_preenche():
    saida = pd.DataFrame([_linha_saida(1, "2026-08-03", 100.0)])
    auditoria = pd.DataFrame([_linha_auditoria(1, "2026-08-04", 100.0, "EXATO_1X1", "X")])

    resultados = processar(saida, auditoria)
    assert resultados[0].preenchido is False
    assert "DATA" in resultados[0].motivo


def test_divergencia_de_valor_nao_preenche():
    saida = pd.DataFrame([_linha_saida(1, "2026-08-03", 100.00)])
    auditoria = pd.DataFrame([_linha_auditoria(1, "2026-08-03", 100.50, "EXATO_1X1", "X")])

    resultados = processar(saida, auditoria)
    assert resultados[0].preenchido is False
    assert "VALOR" in resultados[0].motivo


def test_diferenca_de_1_centavo_e_tolerada():
    saida = pd.DataFrame([_linha_saida(1, "2026-08-03", 100.00)])
    auditoria = pd.DataFrame([_linha_auditoria(1, "2026-08-03", 100.01, "EXATO_1X1", "X")])

    resultados = processar(saida, auditoria)
    assert resultados[0].preenchido is True


def test_status_ambiguo_nao_preenche():
    saida = pd.DataFrame([_linha_saida(1, "2026-08-03", 100.0)])
    auditoria = pd.DataFrame([_linha_auditoria(1, "2026-08-03", 100.0, "AMBIGUO", "X OU Y")])

    resultados = processar(saida, auditoria)
    assert resultados[0].preenchido is False
    assert "STATUS" in resultados[0].motivo


def test_status_nao_encontrado_nao_preenche():
    saida = pd.DataFrame([_linha_saida(1, "2026-08-03", 100.0)])
    auditoria = pd.DataFrame([_linha_auditoria(1, "2026-08-03", 100.0, "NAO_ENCONTRADO", "")])

    resultados = processar(saida, auditoria)
    assert resultados[0].preenchido is False


def test_sequencia_duplicada_na_auditoria_levanta_erro():
    saida = pd.DataFrame([_linha_saida(1, "2026-08-03", 100.0)])
    auditoria = pd.DataFrame(
        [
            _linha_auditoria(1, "2026-08-03", 100.0, "EXATO_1X1", "X"),
            _linha_auditoria(1, "2026-08-03", 100.0, "EXATO_1X1", "Y"),
        ]
    )
    with pytest.raises(MergeSispagError):
        processar(saida, auditoria)
