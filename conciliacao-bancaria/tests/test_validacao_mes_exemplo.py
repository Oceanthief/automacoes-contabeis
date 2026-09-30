"""Validação controlada do motor de conciliação usando o mês de exemplo
(agosto/2026 fictício, gerado por gerar_dados_exemplo.py).

Diferente de tests/test_motor_sispag.py (que usa dados sintéticos para
testar regras isoladas), este arquivo roda o pipeline inteiro contra os
arquivos de exemplo e confirma relações conhecidas de antemão, porque o
cenário foi montado para exercitá-las. No projeto original, este mesmo
teste rodava contra o mês real, conferido manualmente pela contabilidade.
"""
import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from conciliacao.motor_sispag import (
    STATUS_AMBIGUO,
    STATUS_EXATO_1X1,
    STATUS_EXATO_SOMA,
    STATUS_NAO_ENCONTRADO,
    conciliar,
)
from leitura.contas_pagas import ler_contas_pagas, montar_pagamentos
from leitura.extrato_padronizado import ler_extrato_padronizado, listar_sispag_fornecedores
from normalizacao import valor_para_decimal

EXTRATO_PADRONIZADO = RAIZ / "dados_exemplo" / "intermediarios" / "extrato_padronizado.xlsx"
CONTAS_PAGAS = RAIZ / "dados_exemplo" / "contas_pagas_2026_08.xlsx"


@pytest.fixture(scope="module")
def resultados():
    extrato = ler_extrato_padronizado(EXTRATO_PADRONIZADO)
    sispag_df = listar_sispag_fornecedores(extrato).sort_values("SEQUENCIA")
    sispags = [
        (int(l["SEQUENCIA"]), l["DATA"].date(), valor_para_decimal(l["VALOR"]))
        for _, l in sispag_df.iterrows()
    ]
    contas_pagas_df = ler_contas_pagas(CONTAS_PAGAS)
    candidatos, _fora_do_escopo = montar_pagamentos(contas_pagas_df)
    return conciliar(sispags, candidatos)


def _resultado_por_data_e_valor(resultados, data_str: str, valor_str: str):
    dia, mes, ano = (int(x) for x in data_str.split("/"))
    alvo_data = date(ano, mes, dia)
    alvo_valor = Decimal(valor_str)
    encontrados = [r for r in resultados if r.data == alvo_data and r.valor_sispag == alvo_valor]
    assert len(encontrados) == 1, (
        f"Esperava exatamente 1 SISPAG em {data_str} de R$ {valor_str}, achou {len(encontrados)}"
    )
    return encontrados[0]


# ---------------------------------------------------------------------------
# Casos conhecidos do mês de exemplo
# ---------------------------------------------------------------------------


def test_caso_03_08_um_unico_pagamento(resultados):
    r = _resultado_por_data_e_valor(resultados, "03/08/2026", "444.45")
    assert r.status == STATUS_EXATO_1X1


def test_caso_05_08_soma_de_dois_pagamentos(resultados):
    r = _resultado_por_data_e_valor(resultados, "05/08/2026", "2611.83")
    assert r.status == STATUS_EXATO_SOMA
    valores = {p.valor_classificado for p in r.pagamentos}
    assert valores == {Decimal("1316.25"), Decimal("1295.58")}


def test_caso_10_08_soma_de_tres_pagamentos(resultados):
    r = _resultado_por_data_e_valor(resultados, "10/08/2026", "2597.83")
    assert r.status == STATUS_EXATO_SOMA
    valores = {p.valor_classificado for p in r.pagamentos}
    assert valores == {Decimal("1514.29"), Decimal("675.69"), Decimal("407.85")}


def test_caso_18_08_soma_de_dois_pagamentos(resultados):
    r = _resultado_por_data_e_valor(resultados, "18/08/2026", "3315.23")
    assert r.status == STATUS_EXATO_SOMA
    valores = {p.valor_classificado for p in r.pagamentos}
    assert valores == {Decimal("1919.09"), Decimal("1396.14")}


def test_caso_26_08_soma_de_dois_pagamentos(resultados):
    r = _resultado_por_data_e_valor(resultados, "26/08/2026", "3732.50")
    assert r.status == STATUS_EXATO_SOMA
    valores = {p.valor_classificado for p in r.pagamentos}
    assert valores == {Decimal("1885.22"), Decimal("1847.28")}


def test_caso_28_08_titulo_com_duas_classificacoes(resultados):
    """124,41 (juros) + 923,99 (parcelamento) = 1.048,40 num único pagamento."""
    r = _resultado_por_data_e_valor(resultados, "28/08/2026", "1048.40")
    assert r.status == STATUS_EXATO_1X1
    assert len(r.pagamentos[0].classificacoes) == 2


# ---------------------------------------------------------------------------
# Validações gerais sobre os 10 SISPAG
# ---------------------------------------------------------------------------

def test_todos_os_10_sispag_sao_processados(resultados):
    assert len(resultados) == 10


def test_nenhum_pagamento_reutilizado_entre_sispags_diferentes(resultados):
    titulos_usados = [p.titulo for r in resultados for p in r.pagamentos]
    assert len(titulos_usados) == len(set(titulos_usados))


def test_nenhuma_diferenca_monetaria_fora_da_tolerancia(resultados):
    for r in resultados:
        assert abs(r.diferenca) <= Decimal("0.01")


def test_zero_casos_ambiguos(resultados):
    ambiguos = [r for r in resultados if r.status == STATUS_AMBIGUO]
    assert ambiguos == []


def test_zero_casos_nao_encontrados(resultados):
    nao_encontrados = [r for r in resultados if r.status == STATUS_NAO_ENCONTRADO]
    assert nao_encontrados == []


def test_soma_geral_bate_com_total_de_sispag_do_extrato(resultados):
    total_sispag = sum((r.valor_sispag for r in resultados), start=Decimal("0.00"))
    total_encontrado = sum((r.total_encontrado for r in resultados), start=Decimal("0.00"))
    assert total_sispag == Decimal("19658.61")
    assert total_sispag == total_encontrado


def test_matches_individuais_batem_exatamente_com_valor_classificado(resultados):
    for r in resultados:
        if r.status == STATUS_EXATO_1X1:
            assert r.pagamentos[0].valor_classificado == r.valor_sispag


def test_titulos_com_multiplas_classificacoes_nao_geram_duplicidade(resultados):
    """Confirma que títulos com 2+ linhas de VALOR_CLASS na origem (ex.:
    900118, que tem a divergência conhecida de R$ 3,64 com o VALOR
    original) aparecem no máximo 1 vez entre os pagamentos usados, com o
    valor correto (soma de VALOR_CLASS, não o VALOR bruto)."""
    titulos_conhecidos_com_multipla_classificacao = {900114, 900118}
    titulos_usados = [p.titulo for r in resultados for p in r.pagamentos]

    for titulo in titulos_conhecidos_com_multipla_classificacao:
        assert titulos_usados.count(titulo) <= 1

    pagamento_900118 = next(
        (p for r in resultados for p in r.pagamentos if p.titulo == 900118), None
    )
    if pagamento_900118 is not None:
        # soma(VALOR_CLASS) = 162.35, e não o VALOR original de 158.71
        assert pagamento_900118.valor_classificado == Decimal("162.35")
