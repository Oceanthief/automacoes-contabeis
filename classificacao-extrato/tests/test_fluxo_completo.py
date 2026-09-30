"""Testes de ponta a ponta: leitura do .xls, preparar (pendências) e finalizar."""
import csv
import shutil
import sys
from pathlib import Path

import openpyxl
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(RAIZ / "src"), str(RAIZ / "config")]

from exportar_txt import gerar_txt_importacao  # noqa: E402
from fluxo_mensal import processar_mes  # noqa: E402
from lancamentos_finais import gerar_planilha_final  # noqa: E402
from leitura_extrato import ler_extrato  # noqa: E402
from pendencias import nome_arquivo_pendencias  # noqa: E402

DADOS = RAIZ / "dados_exemplo"
EXTRATO = DADOS / "extrato_2026_08.xls"
CONTROLE = DADOS / "controle_fornecedores_clientes.xlsx"
PENDENCIAS_PREENCHIDAS = DADOS / "pendencias_preenchidas_2026_08.xlsx"


# ------------------------------------------------------------------ leitura
def test_leitura_encontra_todos_os_movimentos():
    leitura = ler_extrato(EXTRATO)
    assert leitura.resumo.total_movimentos == 24
    assert leitura.resumo.total_movimentos_sinalizados == 0


def test_saldo_anterior_mais_movimentos_fecha_com_saldo_final():
    assert ler_extrato(EXTRATO).resumo.fecha_com_saldo is True


def test_leitura_nunca_altera_o_arquivo_original():
    import hashlib
    antes = hashlib.md5(EXTRATO.read_bytes()).hexdigest()
    ler_extrato(EXTRATO)
    assert hashlib.md5(EXTRATO.read_bytes()).hexdigest() == antes


# ------------------------------------------------------------------ preparar
def test_preparar_bloqueia_e_lista_as_3_pendencias(tmp_path):
    r = processar_mes(EXTRATO, CONTROLE, 2026, 8, tmp_path, silencioso=True)
    assert r.bloqueado
    assert r.contagem_status == {"OK": 21, "REGRA_PENDENTE": 1, "PENDENTE_MANUAL": 1, "NAO_ENCONTRADO": 1}
    ws = openpyxl.load_workbook(tmp_path / nome_arquivo_pendencias(2026, 8))["Pendencias"]
    assert ws.max_row - 1 == 3


def test_finalizar_sem_resolver_pendencias_nao_gera_nada(tmp_path):
    processar_mes(EXTRATO, CONTROLE, 2026, 8, tmp_path, silencioso=True)
    final = gerar_planilha_final(EXTRATO, CONTROLE, 2026, 8, tmp_path)
    assert not final.gerado
    assert not list(tmp_path.glob("lancamentos_*.xlsx"))


# ------------------------------------------------------------------ finalizar
@pytest.fixture
def mes_finalizado(tmp_path):
    shutil.copy(PENDENCIAS_PREENCHIDAS, tmp_path / nome_arquivo_pendencias(2026, 8))
    final = gerar_planilha_final(EXTRATO, CONTROLE, 2026, 8, tmp_path)
    txt = gerar_txt_importacao(final.caminho_lancamentos, 2026, 8, tmp_path)
    return final, txt


def test_finalizar_com_pendencias_resolvidas_gera_planilha_e_txt(mes_finalizado):
    final, txt = mes_finalizado
    assert final.gerado and txt.gerado
    assert len(final.linhas) == 24
    assert round(sum(l.valor for l in final.linhas), 2) == 65296.35


def test_txt_tem_uma_linha_por_lancamento_e_nenhuma_conta_vazia(mes_finalizado):
    final, txt = mes_finalizado
    with open(txt.caminho_txt, encoding="utf-8-sig", newline="") as f:
        linhas = list(csv.reader(f))
    assert len(linhas) == len(final.linhas)
    for campos in linhas:
        assert all(c.strip() for c in campos[1:3]), campos  # débito e crédito preenchidos


def test_ajuste_manual_chega_ao_lancamento_final(mes_finalizado):
    final, _ = mes_finalizado
    ajustado = [l for l in final.linhas if "FORNECEDOR NOVO SEM CADASTRO" in l.historico_original]
    assert len(ajustado) == 1
    assert ajustado[0].conta_debito == "21105"
