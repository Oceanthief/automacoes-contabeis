"""Testes de granularizacao/reconstrucao_sispag.py com o mês de exemplo."""
import sys
from decimal import Decimal
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from granularizacao.reconstrucao_sispag import (
    EstruturaSispagInsuficienteError,
    reconstruir_resultados_sispag,
)

EXTRATO_PADRONIZADO = RAIZ / "dados_exemplo" / "intermediarios" / "extrato_padronizado.xlsx"
CONTAS_PAGAS = RAIZ / "dados_exemplo" / "contas_pagas_2026_08.xlsx"
AUDITORIA_SISPAG = RAIZ / "dados_exemplo" / "intermediarios" / "auditoria_sispag.xlsx"


def test_reconstroi_10_resultados_e_bate_com_auditoria_publicada():
    resultados = reconstruir_resultados_sispag(
        str(EXTRATO_PADRONIZADO), str(CONTAS_PAGAS), str(AUDITORIA_SISPAG)
    )
    assert len(resultados) == 10


def test_pagamentos_tem_estrutura_completa_nao_e_string_concatenada():
    resultados = reconstruir_resultados_sispag(
        str(EXTRATO_PADRONIZADO), str(CONTAS_PAGAS), str(AUDITORIA_SISPAG)
    )
    algum_com_pagamentos = next(r for r in resultados if r.pagamentos)
    pagamento = algum_com_pagamentos.pagamentos[0]

    assert isinstance(pagamento.titulo, int)
    assert isinstance(pagamento.valor_classificado, Decimal)
    assert isinstance(pagamento.nome, str)


def test_caso_conhecido_soma_de_dois_pagamentos():
    resultados = reconstruir_resultados_sispag(
        str(EXTRATO_PADRONIZADO), str(CONTAS_PAGAS), str(AUDITORIA_SISPAG)
    )
    resultado = next(r for r in resultados if r.sequencia == 10)
    valores = {p.valor_classificado for p in resultado.pagamentos}
    assert valores == {Decimal("1316.25"), Decimal("1295.58")}
