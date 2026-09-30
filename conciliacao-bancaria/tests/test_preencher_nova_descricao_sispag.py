"""Teste de ponta a ponta de preencher_nova_descricao_sispag.py com os
arquivos do mês de exemplo (saida_enriquecida.xlsx + auditoria_sispag.xlsx)."""
import sys
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from preencher_nova_descricao_sispag import PreenchimentoSispagError, executar
from leitura.extrato_padronizado import ler_extrato_padronizado

SAIDA_ENRIQUECIDA = RAIZ / "dados_exemplo" / "intermediarios" / "saida_enriquecida.xlsx"
AUDITORIA = RAIZ / "dados_exemplo" / "intermediarios" / "auditoria_sispag.xlsx"


def test_todos_os_10_sispag_sao_preenchidos(tmp_path):
    destino = tmp_path / "saida_enriquecida_sispag.xlsx"
    resultados = executar(str(SAIDA_ENRIQUECIDA), str(AUDITORIA), str(destino))

    assert len(resultados) == 10
    assert all(r.preenchido for r in resultados)

    saida = pd.read_excel(destino, sheet_name=0)
    sispag = saida[saida["DESCRICAO"] == "SISPAG FORNECEDORES"]
    assert len(sispag) == 10
    assert not sispag["NOVA DESCRICAO"].isna().any()
    assert not (sispag["NOVA DESCRICAO"] == "").any()


def test_contagem_5_exato_1x1_e_5_exato_soma(tmp_path):
    destino = tmp_path / "saida_enriquecida_sispag.xlsx"
    executar(str(SAIDA_ENRIQUECIDA), str(AUDITORIA), str(destino))

    auditoria_df = pd.read_excel(AUDITORIA, sheet_name=0)
    contagem = auditoria_df["STATUS"].value_counts()
    assert contagem["EXATO_1X1"] == 5
    assert contagem["EXATO_SOMA"] == 5


def test_linhas_nao_sispag_permanecem_intactas(tmp_path):
    """Usa o LEITOR COMPARTILHADO (leitura.extrato_padronizado) em vez de
    pd.read_excel cru dos dois lados — comparar "antes" e "depois" com o
    mesmo bug de leitura nos dois lados esconderia uma corrupção real (foi
    exatamente o que aconteceu: os dois ficavam iguais a "248.0", então o
    teste passava mesmo com o zero à esquerda perdido)."""
    destino = tmp_path / "saida_enriquecida_sispag.xlsx"
    executar(str(SAIDA_ENRIQUECIDA), str(AUDITORIA), str(destino))

    antes = ler_extrato_padronizado(SAIDA_ENRIQUECIDA)
    depois = ler_extrato_padronizado(destino)

    for col in ["DEBITO", "CREDITO", "DATA", "VALOR", "SEQUENCIA"]:
        assert (antes[col].astype(str).fillna("") == depois[col].astype(str).fillna("")).all()

    nao_sispag_antes = antes[antes["DESCRICAO"] != "SISPAG FORNECEDORES"]["NOVA DESCRICAO"].fillna("")
    nao_sispag_depois = depois[depois["DESCRICAO"] != "SISPAG FORNECEDORES"]["NOVA DESCRICAO"].fillna("")
    assert list(nao_sispag_antes) == list(nao_sispag_depois)

    # trava específica do bug já detectado: conta de tarifa não pode perder o zero
    tarifa = depois[depois["DESCRICAO"].str.startswith("TAR ", na=False)].iloc[0]
    assert tarifa["DEBITO"] == "00248"


def test_fornecedor_repetido_deduplicado_no_sispag_de_seq_58(tmp_path):
    destino = tmp_path / "saida_enriquecida_sispag.xlsx"
    executar(str(SAIDA_ENRIQUECIDA), str(AUDITORIA), str(destino))

    saida = pd.read_excel(destino, sheet_name=0)
    linha = saida[saida["SEQUENCIA"] == 58].iloc[0]
    assert linha["NOVA DESCRICAO"] == "ENERGIA REGIONAL DISTRIBUIDORA SA"
    assert ";" not in linha["NOVA DESCRICAO"]  # não pode sobrar duplicata


def test_erro_se_saida_igual_a_entrada():
    with pytest.raises(PreenchimentoSispagError):
        executar(str(SAIDA_ENRIQUECIDA), str(AUDITORIA), str(SAIDA_ENRIQUECIDA))
