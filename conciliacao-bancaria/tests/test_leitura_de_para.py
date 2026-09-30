"""Testes de leitura/de_para.py."""
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from leitura.de_para import DeParaError, ler_de_para

DE_PARA = RAIZ / "dados_exemplo" / "de_para_contabil.xlsx"


def test_le_regras_e_remove_linha_em_branco():
    df = ler_de_para(DE_PARA)
    # 20 linhas na planilha, 1 delas totalmente em branco (separador)
    assert len(df) == 19
    assert df["CHAVE_MATCH"].notna().all()


def test_conta_contabil_preserva_zero_a_esquerda():
    df = ler_de_para(DE_PARA)
    contas_com_zero = df[df["CONTA_CONTABIL"].astype(str).str.startswith("0")]
    assert len(contas_com_zero) > 0


def test_erro_se_arquivo_nao_existe():
    with pytest.raises(DeParaError):
        ler_de_para("nao_existe.xlsx")
