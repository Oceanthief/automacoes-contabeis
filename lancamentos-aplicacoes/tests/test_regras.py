"""Testes da regra contábil do CONTAMAX e do layout do TXT de importação."""
import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

import main  # noqa: E402
from contamax.models import Movimentacao, StatusLancamento  # noqa: E402
from contamax.regras import (  # noqa: E402
    CONTA_APLICACAO_CONTAMAX,
    CONTA_SALDO_CORRENTE,
    HISTORICO_APLICACAO,
    HISTORICO_RESGATE,
    agrupar_por_data,
    aplicar_regra_contamax,
)

DIA = date(2026, 8, 10)


# ------------------------------------------------------------------ regra do dia
def test_saiu_mais_do_que_entrou_gera_resgate():
    r = aplicar_regra_contamax(DIA, Decimal("1000.00"), Decimal("429.14"))
    assert r.status == StatusLancamento.GERADO
    assert (r.conta_debito, r.conta_credito) == (CONTA_APLICACAO_CONTAMAX, CONTA_SALDO_CORRENTE)
    assert r.historico == HISTORICO_RESGATE
    assert r.valor == Decimal("570.86")


def test_entrou_mais_do_que_saiu_gera_aplicacao_com_valor_positivo():
    r = aplicar_regra_contamax(DIA, Decimal("100.00"), Decimal("350.50"))
    assert (r.conta_debito, r.conta_credito) == (CONTA_SALDO_CORRENTE, CONTA_APLICACAO_CONTAMAX)
    assert r.historico == HISTORICO_APLICACAO
    assert r.valor == Decimal("250.50")


def test_empate_nao_gera_lancamento():
    """Decisão 5.1 da especificação: débito = crédito não gera linha no TXT."""
    r = aplicar_regra_contamax(DIA, Decimal("1200.00"), Decimal("1200.00"))
    assert r.status == StatusLancamento.SEM_LANCAMENTO_EMPATE
    assert r.valor is None and r.conta_debito is None


# ------------------------------------------------------------------ agrupamento
def _mov(dia, debito="0.00", credito="0.00"):
    return Movimentacao(date(2026, 8, dia), Decimal(debito), Decimal(credito))


def test_agrupamento_soma_varios_movimentos_do_mesmo_dia():
    totais = agrupar_por_data([_mov(10, debito="640.30"), _mov(10, debito="215.00"),
                               _mov(10, debito="49.90"), _mov(10, credito="150.00"), _mov(11, credito="1")])
    assert totais[date(2026, 8, 10)] == (Decimal("905.20"), Decimal("150.00"))
    assert len(totais) == 2


def test_soma_em_decimal_nao_tem_erro_de_ponto_flutuante():
    totais = agrupar_por_data([_mov(10, debito="0.10"), _mov(10, debito="0.20")])
    assert totais[date(2026, 8, 10)][0] == Decimal("0.30")  # em float: 0.30000000000000004


# ------------------------------------------------------------------ TXT (mês de exemplo)
@pytest.fixture
def linhas_txt(tmp_path, monkeypatch):
    monkeypatch.setattr(main, "DIRETORIO_SAIDA_BASE", tmp_path)
    assert main.main([str(RAIZ / "dados_exemplo" / "extrato_2026_08.xls")]) == 0
    bruto = (tmp_path / "2026-08" / "contamax_2026-08.txt").read_bytes()
    assert bruto.startswith(b"\xef\xbb\xbf"), "TXT precisa ter BOM (UTF-8-sig)"
    return bruto.decode("utf-8-sig").split("\r\n")[:-1]


def test_linha_no_layout_confirmado_pela_contabilidade(linhas_txt):
    assert linhas_txt[9] == '"","11023","11145","25/08/2026","570.86","00013","VLR REF. RESGATE CONTAMAX","10"'


def test_dia_de_empate_nao_aparece_no_txt(linhas_txt):
    assert not any('"12/08/2026"' in linha for linha in linhas_txt)
    assert len(linhas_txt) == 11


def test_sequencia_comeca_em_1_sem_zero_a_esquerda(linhas_txt):
    sequencias = [linha.rsplit(",", 1)[1].strip('"') for linha in linhas_txt]
    assert sequencias == [str(i) for i in range(1, 12)]
