"""Testes de ponta a ponta do enriquecimento do extrato padronizado."""
import sys
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from enriquecer_extrato import EnriquecimentoError, executar
from enriquecimento.classificacao_extrato import CONTA_TARIFA
from leitura.extrato_padronizado import ler_extrato_padronizado

EXTRATO_PADRONIZADO = RAIZ / "dados_exemplo" / "intermediarios" / "extrato_padronizado.xlsx"


def test_enriquecimento_completo_do_mes_de_exemplo(tmp_path):
    destino = tmp_path / "saida_enriquecida.xlsx"
    linhas = executar(str(EXTRATO_PADRONIZADO), str(destino))

    # Critério 1: 58 movimentações continuam existindo
    assert len(linhas) == 58

    # Critério 4: toda tarifa tem DEBITO=00248 e CREDITO=10001
    tarifas = [l for l in linhas if l["_eh_tarifa"]]
    assert len(tarifas) == 6
    assert all(l["DEBITO"] == CONTA_TARIFA and l["CREDITO"] == "10001" for l in tarifas)

    # Critério 6: SISPAG FORNECEDORES com NOVA DESCRICAO vazia
    sispags = [l for l in linhas if l["DESCRICAO"] == "SISPAG FORNECEDORES"]
    assert len(sispags) == 10
    assert all(l["NOVA DESCRICAO"] == "" for l in sispags)

    # Critério 5: BO EST LANC TARIFA não vira despesa de tarifa
    bo_tarifa = [l for l in linhas if "LANC TARIFA" in l["DESCRICAO"]]
    assert len(bo_tarifa) == 1
    assert bo_tarifa[0]["_eh_tarifa"] is False

    assert destino.exists()


def test_sequencia_preservada_e_arquivo_gerado_tem_9_colunas(tmp_path):
    destino = tmp_path / "saida_enriquecida.xlsx"
    executar(str(EXTRATO_PADRONIZADO), str(destino))

    saida = pd.read_excel(destino, sheet_name=0)
    original = pd.read_excel(EXTRATO_PADRONIZADO, sheet_name=0)

    assert list(saida["SEQUENCIA"]) == list(original["SEQUENCIA"])
    assert len(saida.columns) == 9
    assert "NOVA DESCRICAO" in saida.columns


def test_descricao_original_intacta_na_saida(tmp_path):
    destino = tmp_path / "saida_enriquecida.xlsx"
    executar(str(EXTRATO_PADRONIZADO), str(destino))

    saida = pd.read_excel(destino, sheet_name=0)
    original = pd.read_excel(EXTRATO_PADRONIZADO, sheet_name=0)

    assert list(saida["DESCRICAO"]) == list(original["DESCRICAO"])


def test_conta_de_tarifa_preserva_zero_a_esquerda(tmp_path):
    """Usa o LEITOR COMPARTILHADO do projeto (não um dtype forçado manualmente
    só no teste) — é o caminho real que qualquer fase seguinte vai usar para
    reabrir este arquivo. Detectou um bug real: sem dtype=str no leitor, o
    pandas inferia DEBITO como float64 e '00248' virava '248.0'."""
    destino = tmp_path / "saida_enriquecida.xlsx"
    executar(str(EXTRATO_PADRONIZADO), str(destino))

    saida = ler_extrato_padronizado(destino)
    tarifa = saida[saida["DESCRICAO"].str.startswith("TAR ", na=False)].iloc[0]
    assert tarifa["DEBITO"] == "00248"  # não "248" nem "248.0"


def test_erro_se_saida_igual_a_entrada():
    with pytest.raises(EnriquecimentoError):
        executar(str(EXTRATO_PADRONIZADO), str(EXTRATO_PADRONIZADO))
