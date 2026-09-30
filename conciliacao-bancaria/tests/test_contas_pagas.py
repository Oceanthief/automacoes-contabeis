"""Testes do módulo de normalização de Contas Pagas."""
import sys
from decimal import Decimal
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from leitura.contas_pagas import (
    CONTA_ALVO,
    ContasPagasError,
    ler_contas_pagas,
    listar_divergencias,
    montar_pagamentos,
)
from normalizacao import normalizar_data

ARQUIVO_AMOSTRA = RAIZ / "dados_exemplo" / "contas_pagas_2026_08.xlsx"


# ---------------------------------------------------------------------------
# Critérios de aceite com o mês de exemplo (gerar_dados_exemplo.py)
# ---------------------------------------------------------------------------

def test_quantidade_de_linhas_de_origem():
    df = ler_contas_pagas(ARQUIVO_AMOSTRA)
    assert len(df) == 20


def test_titulo_com_duas_classificacoes_vira_um_unico_pagamento():
    df = ler_contas_pagas(ARQUIVO_AMOSTRA)
    candidatos, fora_do_escopo = montar_pagamentos(df)
    todos = candidatos + fora_do_escopo

    pagamento = next(p for p in todos if p.titulo == 900114)
    assert len(pagamento.classificacoes) == 2
    assert pagamento.valor_classificado == Decimal("1048.40")
    # confirma que é só 1 pagamento na lista, não 2
    ocorrencias = [p for p in todos if p.titulo == 900114]
    assert len(ocorrencias) == 1


def test_nenhum_titulo_duplicado_na_lista_final():
    df = ler_contas_pagas(ARQUIVO_AMOSTRA)
    candidatos, fora_do_escopo = montar_pagamentos(df)
    todos = candidatos + fora_do_escopo
    chaves = [(p.titulo, p.data_efetiva) for p in todos]
    assert len(chaves) == len(set(chaves))


def test_quantidade_de_pagamentos_unicos():
    df = ler_contas_pagas(ARQUIVO_AMOSTRA)
    candidatos, fora_do_escopo = montar_pagamentos(df)
    # 18 TITULOs únicos no mês de exemplo (20 linhas, 2 títulos com 2 classificações)
    assert len(candidatos) + len(fora_do_escopo) == 18


def test_filtro_conta_alvo():
    df = ler_contas_pagas(ARQUIVO_AMOSTRA)
    candidatos, fora_do_escopo = montar_pagamentos(df)
    assert all(p.conta == CONTA_ALVO for p in candidatos)
    assert all(p.conta != CONTA_ALVO for p in fora_do_escopo)
    # o mês de exemplo tem 2 pagamentos no "341 - Cartão Corporativo"
    assert len(fora_do_escopo) >= 1


def test_divergencia_encontrada_no_titulo_900118():
    df = ler_contas_pagas(ARQUIVO_AMOSTRA)
    candidatos, fora_do_escopo = montar_pagamentos(df)
    divergentes = listar_divergencias(candidatos + fora_do_escopo)

    titulos_divergentes = {p.titulo for p in divergentes}
    assert 900118 in titulos_divergentes

    pagamento = next(p for p in divergentes if p.titulo == 900118)
    assert pagamento.valor_original == Decimal("158.71")
    assert pagamento.valor_classificado == Decimal("162.35")
    assert pagamento.divergencia == Decimal("-3.64")


def test_divergencia_respeita_tolerancia_de_1_centavo():
    df = ler_contas_pagas(ARQUIVO_AMOSTRA)
    candidatos, fora_do_escopo = montar_pagamentos(df)
    todos = candidatos + fora_do_escopo
    # nenhum pagamento com diferença <= 0,01 deve aparecer na lista de divergências
    divergentes = listar_divergencias(todos)
    for p in divergentes:
        assert abs(p.divergencia) > Decimal("0.01")


# ---------------------------------------------------------------------------
# Critério de aceite explícito da tarefa: 124,41 + 923,99 = 1048,40
# ---------------------------------------------------------------------------

def test_caso_especifico_soma_classificacoes():
    df = pd.DataFrame(
        [
            {
                "TITULO": 999999, "NOME": "Fornecedor Teste", "NF_NUM": "NF1",
                "DATA_EFETIVA_CH": "28/08/2026", "VALOR": 1048.40, "VALOR_CLASS": 124.41,
                "CLASSIFIC_COD": "2.10.053", "CLASSIFIC_DESC": "Juros", "HISTORICO": "Teste",
                "FORMA_PGTO": "Cobrança Bancária", "CONTA": CONTA_ALVO,
            },
            {
                "TITULO": 999999, "NOME": "Fornecedor Teste", "NF_NUM": "NF1",
                "DATA_EFETIVA_CH": "28/08/2026", "VALOR": None, "VALOR_CLASS": 923.99,
                "CLASSIFIC_COD": "2.10.086", "CLASSIFIC_DESC": "Parcela", "HISTORICO": "Teste",
                "FORMA_PGTO": "Cobrança Bancária", "CONTA": CONTA_ALVO,
            },
        ]
    )
    candidatos, fora_do_escopo = montar_pagamentos(df)
    assert len(candidatos) == 1
    assert len(fora_do_escopo) == 0
    assert candidatos[0].valor_classificado == Decimal("1048.40")
    assert not candidatos[0].tem_divergencia_relevante


# ---------------------------------------------------------------------------
# normalizar_data: robustez contra os tipos de entrada exigidos
# ---------------------------------------------------------------------------

def test_normalizar_data_datetime():
    import datetime as dt
    assert normalizar_data(dt.datetime(2026, 8, 28, 13, 45)) == dt.date(2026, 8, 28)


def test_normalizar_data_date():
    import datetime as dt
    assert normalizar_data(dt.date(2026, 8, 28)) == dt.date(2026, 8, 28)


def test_normalizar_data_texto_brasileiro():
    import datetime as dt
    assert normalizar_data("28/08/2026") == dt.date(2026, 8, 28)


def test_normalizar_data_texto_iso():
    import datetime as dt
    assert normalizar_data("2026-08-28") == dt.date(2026, 8, 28)


def test_normalizar_data_serial_excel():
    import datetime as dt
    # 46262 = 28/08/2026 no sistema de serial do Excel (dias desde 1899-12-30)
    assert normalizar_data(46262) == dt.date(2026, 8, 28)


def test_normalizar_data_nula_levanta_erro():
    with pytest.raises(ValueError):
        normalizar_data(None)
    with pytest.raises(ValueError):
        normalizar_data(float("nan"))


# ---------------------------------------------------------------------------
# Integridade: campos invariantes não podem variar dentro de um pagamento
# ---------------------------------------------------------------------------

def test_campo_invariante_divergente_levanta_erro():
    df = pd.DataFrame(
        [
            {
                "TITULO": 1, "NOME": "Fornecedor A", "NF_NUM": "1",
                "DATA_EFETIVA_CH": "10/08/2026", "VALOR": 100.0, "VALOR_CLASS": 50.0,
                "CLASSIFIC_COD": "x", "CLASSIFIC_DESC": "x", "HISTORICO": "x",
                "FORMA_PGTO": "Pix/Transferência", "CONTA": CONTA_ALVO,
            },
            {
                # mesmo TITULO e data, mas NOME diferente -> deveria ser impossível
                "TITULO": 1, "NOME": "Fornecedor B", "NF_NUM": "1",
                "DATA_EFETIVA_CH": "10/08/2026", "VALOR": None, "VALOR_CLASS": 50.0,
                "CLASSIFIC_COD": "y", "CLASSIFIC_DESC": "y", "HISTORICO": "x",
                "FORMA_PGTO": "Pix/Transferência", "CONTA": CONTA_ALVO,
            },
        ]
    )
    with pytest.raises(ContasPagasError):
        montar_pagamentos(df)


def test_erro_se_arquivo_nao_existe():
    with pytest.raises(ContasPagasError):
        ler_contas_pagas("nao_existe.xlsx")


def test_erro_se_coluna_esperada_faltando():
    df = pd.DataFrame([{"TITULO": 1}])
    with pytest.raises(ContasPagasError):
        # simula arquivo sem as colunas certas, salvando e relendo
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            df.to_excel(tmp.name, index=False)
            ler_contas_pagas(tmp.name)
