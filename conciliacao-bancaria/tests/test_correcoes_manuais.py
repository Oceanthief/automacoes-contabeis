"""Testes de leitura/correcoes_manuais.py."""
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from leitura.correcoes_manuais import CorrecoesManuaisError, ler_correcoes_manuais, normalizar_codigo_conta_manual

CORRECOES = RAIZ / "dados_exemplo" / "correcoes_manuais.xlsx"


def test_le_as_correcoes_do_mes_de_exemplo():
    df = ler_correcoes_manuais(CORRECOES)
    assert len(df) == 1


def test_debito_manual_preserva_zero_a_esquerda():
    df = ler_correcoes_manuais(CORRECOES)
    linha = df[df["SEQUENCIA_CONTABIL"] == 36].iloc[0]
    assert linha["DEBITO_MANUAL"] == "00032"


@pytest.mark.parametrize(
    "entrada,esperado",
    [
        ("32", "00032"),
        ("243", "00243"),
        ("00243", "00243"),
        ("8685", "8685"),  # exceção confirmada pelo usuário
        ("94468", "94468"),
        (None, ""),
        ("", ""),
    ],
)
def test_normalizar_codigo_conta_manual(entrada, esperado):
    assert normalizar_codigo_conta_manual(entrada) == esperado


def test_erro_se_arquivo_nao_existe():
    with pytest.raises(CorrecoesManuaisError):
        ler_correcoes_manuais("nao_existe.xlsx")
