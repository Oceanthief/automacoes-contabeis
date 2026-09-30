"""Testes do motor de conciliação SISPAG FORNECEDORES."""
import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from conciliacao.motor_sispag import (
    STATUS_AMBIGUO,
    STATUS_EXATO_1X1,
    STATUS_EXATO_SOMA,
    STATUS_NAO_ENCONTRADO,
    conciliar,
)
from leitura.contas_pagas import Pagamento


def _pagamento(titulo, valor, data_efetiva=date(2026, 8, 3), nome="Fornecedor", conta="341 - Conta Movimento"):
    return Pagamento(
        titulo=titulo,
        data_efetiva=data_efetiva,
        nome=nome,
        nf_num=None,
        forma_pgto="Cobrança Bancária",
        conta=conta,
        historico="teste",
        valor_original=Decimal(str(valor)),
        valor_classificado=Decimal(str(valor)),
        classificacoes=[],
    )


# ---------------------------------------------------------------------------
# Caso do enunciado: match individual simples
# ---------------------------------------------------------------------------

def test_match_individual_exato():
    pagamentos = [_pagamento(1, "1197.16")]
    sispags = [(1, date(2026, 8, 3), Decimal("1197.16"))]

    resultados = conciliar(sispags, pagamentos)

    assert len(resultados) == 1
    assert resultados[0].status == STATUS_EXATO_1X1
    assert resultados[0].pagamentos[0].titulo == 1
    assert resultados[0].total_encontrado == Decimal("1197.16")
    assert resultados[0].diferenca == Decimal("0.00")


# ---------------------------------------------------------------------------
# Caso do enunciado: combinação de dois pagamentos (720,41 + 476,75 = 1197,16)
# ---------------------------------------------------------------------------

def test_match_por_combinacao_exemplo_do_enunciado():
    pagamentos = [_pagamento(10, "720.41"), _pagamento(20, "476.75")]
    sispags = [(1, date(2026, 8, 3), Decimal("1197.16"))]

    resultados = conciliar(sispags, pagamentos)

    assert resultados[0].status == STATUS_EXATO_SOMA
    titulos = {p.titulo for p in resultados[0].pagamentos}
    assert titulos == {10, 20}
    assert resultados[0].total_encontrado == Decimal("1197.16")


# ---------------------------------------------------------------------------
# Ambiguidade individual: dois pagamentos com o mesmo valor no mesmo dia
# ---------------------------------------------------------------------------

def test_ambiguo_individual_dois_candidatos_mesmo_valor():
    pagamentos = [_pagamento(1, "500.00"), _pagamento(2, "500.00")]
    sispags = [(1, date(2026, 8, 3), Decimal("500.00"))]

    resultados = conciliar(sispags, pagamentos)

    assert resultados[0].status == STATUS_AMBIGUO
    assert len(resultados[0].alternativas) == 2
    # não pode ter consumido nenhum pagamento do pool
    assert resultados[0].pagamentos == []
    assert resultados[0].total_encontrado == Decimal("0.00")


# ---------------------------------------------------------------------------
# Ambiguidade por soma: duas combinações diferentes batem o mesmo alvo
# ---------------------------------------------------------------------------

def test_ambiguo_por_soma_duas_combinacoes_possiveis():
    # 100 + 200 = 300  E  50 + 250 = 300 -> duas combinações válidas
    pagamentos = [
        _pagamento(1, "100.00"),
        _pagamento(2, "200.00"),
        _pagamento(3, "50.00"),
        _pagamento(4, "250.00"),
    ]
    sispags = [(1, date(2026, 8, 3), Decimal("300.00"))]

    resultados = conciliar(sispags, pagamentos)

    assert resultados[0].status == STATUS_AMBIGUO
    assert len(resultados[0].alternativas) == 2
    assert resultados[0].pagamentos == []


# ---------------------------------------------------------------------------
# Nenhuma combinação explica o valor
# ---------------------------------------------------------------------------

def test_nao_encontrado_quando_nenhuma_combinacao_bate():
    pagamentos = [_pagamento(1, "10.00"), _pagamento(2, "20.00")]
    sispags = [(1, date(2026, 8, 3), Decimal("999.99"))]

    resultados = conciliar(sispags, pagamentos)

    assert resultados[0].status == STATUS_NAO_ENCONTRADO
    assert resultados[0].total_encontrado == Decimal("0.00")
    assert resultados[0].diferenca == Decimal("999.99")


# ---------------------------------------------------------------------------
# Um pagamento nunca pode ser usado em dois SISPAG diferentes
# ---------------------------------------------------------------------------

def test_pagamento_nao_e_reutilizado_entre_sispags():
    pagamentos = [_pagamento(1, "1000.00")]
    sispags = [
        (1, date(2026, 8, 3), Decimal("1000.00")),
        (2, date(2026, 8, 3), Decimal("1000.00")),
    ]

    resultados = conciliar(sispags, pagamentos)

    assert resultados[0].status == STATUS_EXATO_1X1
    # o segundo SISPAG não pode reencontrar o mesmo pagamento já usado
    assert resultados[1].status == STATUS_NAO_ENCONTRADO

    titulos_usados = [p.titulo for r in resultados for p in r.pagamentos]
    assert len(titulos_usados) == len(set(titulos_usados))


def test_dois_sispags_iguais_dois_pagamentos_iguais_ficam_ambiguos():
    """Caso legítimo de ambiguidade, não um bug: com 2 pagamentos de R$ 300 e
    2 SISPAG de R$ 300 no mesmo dia, o primeiro SISPAG já enxerga 2
    candidatos válidos (os dois pagamentos) — não há como saber qual dos
    dois é o certo, então a regra manda marcar AMBIGUO em vez de "adivinhar"
    pela ordem de chegada. O segundo SISPAG, com os mesmos 2 candidatos
    ainda disponíveis (nada foi consumido), também fica AMBIGUO.
    """
    pagamentos = [_pagamento(1, "300.00"), _pagamento(2, "300.00")]
    sispags = [
        (1, date(2026, 8, 3), Decimal("300.00")),
        (2, date(2026, 8, 3), Decimal("300.00")),
    ]

    resultados = conciliar(sispags, pagamentos)

    assert resultados[0].status == STATUS_AMBIGUO
    assert resultados[1].status == STATUS_AMBIGUO
    # nenhum pagamento foi consumido — ambos continuam como alternativas
    assert resultados[0].pagamentos == []
    assert resultados[1].pagamentos == []


# ---------------------------------------------------------------------------
# Datas diferentes não se misturam
# ---------------------------------------------------------------------------

def test_pagamento_de_outra_data_nao_e_considerado():
    pagamentos = [_pagamento(1, "1197.16", data_efetiva=date(2026, 8, 4))]
    sispags = [(1, date(2026, 8, 3), Decimal("1197.16"))]

    resultados = conciliar(sispags, pagamentos)

    assert resultados[0].status == STATUS_NAO_ENCONTRADO


# ---------------------------------------------------------------------------
# Poda: candidato maior que o alvo não deve nem entrar na busca de combinação
# ---------------------------------------------------------------------------

def test_candidato_maior_que_alvo_e_descartado_da_combinacao():
    pagamentos = [_pagamento(1, "5000.00"), _pagamento(2, "300.00"), _pagamento(3, "200.00")]
    sispags = [(1, date(2026, 8, 3), Decimal("500.00"))]

    resultados = conciliar(sispags, pagamentos)

    assert resultados[0].status == STATUS_EXATO_SOMA
    titulos = {p.titulo for p in resultados[0].pagamentos}
    assert titulos == {2, 3}  # o de 5000 nunca poderia ter entrado


# ---------------------------------------------------------------------------
# Critério de aceite: soma dos pagamentos ligados bate com o SISPAG (±0,01)
# ---------------------------------------------------------------------------

def test_soma_dos_matches_bate_com_sispag_dentro_da_tolerancia():
    pagamentos = [_pagamento(1, "720.41"), _pagamento(2, "476.75")]
    sispags = [(1, date(2026, 8, 3), Decimal("1197.16"))]

    resultados = conciliar(sispags, pagamentos)
    r = resultados[0]
    assert abs(r.valor_sispag - r.total_encontrado) <= Decimal("0.01")
