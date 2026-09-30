"""Teste de ponta a ponta de granularizar_sispag.executar() com os arquivos
do mês de exemplo — inclui a trava de regressão do bug de zero à esquerda
encontrado na auditoria (CONTA_ENCONTRADA lida sem dtype=str)."""
import sys
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from granularizar_sispag import executar

EXTRATO = RAIZ / "dados_exemplo" / "intermediarios" / "extrato_padronizado.xlsx"
CONTAS_PAGAS = RAIZ / "dados_exemplo" / "contas_pagas_2026_08.xlsx"
AUDITORIA_SISPAG = RAIZ / "dados_exemplo" / "intermediarios" / "auditoria_sispag.xlsx"
AUDITORIA_DE_PARA = RAIZ / "dados_exemplo" / "intermediarios" / "auditoria_de_para.xlsx"
DE_PARA = RAIZ / "dados_exemplo" / "de_para_contabil.xlsx"


pytestmark = pytest.mark.skipif(
    not AUDITORIA_DE_PARA.exists(), reason="auditoria_de_para.xlsx da Fase 4 não está disponível neste ambiente"
)


def test_pipeline_completo_preserva_conta_de_tarifa_com_zero_a_esquerda(tmp_path):
    saida_contabil = tmp_path / "saida_contabil.xlsx"
    auditoria_gran = tmp_path / "auditoria_granularizacao.xlsx"

    executar(
        str(EXTRATO), str(CONTAS_PAGAS), str(AUDITORIA_SISPAG),
        str(AUDITORIA_DE_PARA), str(DE_PARA),
        str(saida_contabil), str(auditoria_gran),
    )

    saida = pd.read_excel(saida_contabil, sheet_name=0, dtype={"DEBITO": str, "CREDITO": str})
    tarifa = saida[saida["DESCRICAO_ORIGINAL"].str.startswith("TAR ", na=False)].iloc[0]
    assert tarifa["DEBITO"] == "00248"  # não "248" nem "248.0"


def test_pipeline_completo_64_linhas_e_nenhuma_duplicidade(tmp_path):
    saida_contabil = tmp_path / "saida_contabil.xlsx"
    auditoria_gran = tmp_path / "auditoria_granularizacao.xlsx"

    executar(
        str(EXTRATO), str(CONTAS_PAGAS), str(AUDITORIA_SISPAG),
        str(AUDITORIA_DE_PARA), str(DE_PARA),
        str(saida_contabil), str(auditoria_gran),
    )

    saida = pd.read_excel(saida_contabil, sheet_name=0)
    assert len(saida) == 64
    assert saida["SEQUENCIA_CONTABIL"].is_unique

    sispag = saida[saida["DESCRICAO_ORIGINAL"] == "SISPAG FORNECEDORES"]
    titulos = sispag["TITULO"].dropna()
    assert titulos.nunique() == len(titulos) == 16
