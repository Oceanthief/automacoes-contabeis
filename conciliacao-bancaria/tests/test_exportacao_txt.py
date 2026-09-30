"""Testes de exportacao_txt.py."""
import sys
from decimal import Decimal
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from exportacao_txt import (
    ExportacaoError,
    STATUS_TXT_VALIDADO,
    escrever_txt,
    executar,
    formatar_linhas,
    montar_linha_txt,
    reler_e_reconstruir,
    validar_pos_geracao,
)

# TXT de referência no layout de importação do ERP (gerado para o mês de exemplo)
ARQUIVO_REFERENCIA = RAIZ / "dados_exemplo" / "intermediarios" / "referencia_importacao_contabil.txt"

RELATORIO_VALIDACAO = RAIZ / "dados_exemplo" / "intermediarios" / "relatorio_validacao.xlsx"


def _linha_aprovada(**overrides):
    base = {
        "SEQUENCIA_CONTABIL": 1, "SEQUENCIA_BANCO": 1, "DEBITO": "00248", "CREDITO": "10001",
        "DATA": pd.Timestamp("2026-08-03"), "VALOR": 100.0,
        "NOVA_DESCRICAO": "VLR REF. PGTO TAR BLOQUETO ITAU", "STATUS_VALIDACAO": "APROVADO",
    }
    base.update(overrides)
    return base


# ---------------------------------------------------------------------------
# montar_linha_txt — bate byte a byte com o arquivo de referência
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not ARQUIVO_REFERENCIA.exists(), reason="Arquivos.txt de referência não disponível")
def test_reconstroi_linhas_reais_do_arquivo_de_referencia():
    with open(ARQUIVO_REFERENCIA, encoding="utf-8-sig", newline="") as f:
        linhas_reais = f.read().split("\r\n")
    linhas_reais = [l for l in linhas_reais if l]

    import csv
    with open(ARQUIVO_REFERENCIA, encoding="utf-8-sig", newline="") as f:
        campos = list(csv.reader(f))

    for i in [0, 3, 4, 5, len(campos) - 1]:
        c = campos[i]
        reconstruida = montar_linha_txt(c[1], c[2], c[3], c[4], c[6], c[7])
        assert reconstruida == linhas_reais[i]


def test_montar_linha_txt_formato_basico():
    linha = montar_linha_txt("00248", "10001", "03/08/2026", "100.00", "VLR REF. PGTO X", 1)
    assert linha == '"","00248","10001","03/08/2026","100.00","","VLR REF. PGTO X","1"'


def test_montar_linha_txt_escapa_aspas_internas():
    linha = montar_linha_txt("00248", "10001", "03/08/2026", "100.00", 'FORNECEDOR "APELIDO"', 1)
    assert '""APELIDO""' in linha


# ---------------------------------------------------------------------------
# formatar_linhas
# ---------------------------------------------------------------------------

def test_formatar_linhas_preserva_conta_com_4_digitos():
    """Caso real do projeto: '8685' é uma exceção confirmada — não pode ser
    forçada para 5 dígitos aqui."""
    df = pd.DataFrame([_linha_aprovada(DEBITO="8685")])
    linhas = formatar_linhas(df)
    assert linhas[0].debito == "8685"


def test_formatar_linhas_data_dd_mm_aaaa():
    df = pd.DataFrame([_linha_aprovada(DATA=pd.Timestamp("2026-01-05"))])
    linhas = formatar_linhas(df)
    assert "05/01/2026" in linhas[0].linha_txt


def test_formatar_linhas_valor_duas_casas():
    df = pd.DataFrame([_linha_aprovada(VALOR=9355.6)])
    linhas = formatar_linhas(df)
    assert "9355.60" in linhas[0].linha_txt


def test_formatar_linhas_conta_vazia_levanta_erro():
    df = pd.DataFrame([_linha_aprovada(DEBITO="")])
    with pytest.raises(ExportacaoError):
        formatar_linhas(df)


# ---------------------------------------------------------------------------
# Escrita + releitura + round-trip
# ---------------------------------------------------------------------------

def test_escrever_e_reler_bate_exatamente(tmp_path):
    df = pd.DataFrame(
        [
            _linha_aprovada(SEQUENCIA_CONTABIL=1, SEQUENCIA_BANCO=1),
            _linha_aprovada(SEQUENCIA_CONTABIL=2, SEQUENCIA_BANCO=2, DEBITO="10001", CREDITO="05620", VALOR=50.0),
        ]
    )
    linhas = formatar_linhas(df)
    caminho = tmp_path / "saida.txt"
    escrever_txt(linhas, caminho)

    with open(caminho, "rb") as f:
        conteudo = f.read()
    assert conteudo.startswith(b"\xef\xbb\xbf")
    assert conteudo.endswith(b"\r\n")

    problemas = validar_pos_geracao(linhas, caminho)
    assert problemas == []


def test_round_trip_detecta_linha_faltante(tmp_path):
    df = pd.DataFrame([_linha_aprovada(SEQUENCIA_CONTABIL=1), _linha_aprovada(SEQUENCIA_CONTABIL=2)])
    linhas = formatar_linhas(df)
    caminho = tmp_path / "saida.txt"
    escrever_txt(linhas[:1], caminho)  # escreve só 1 das 2 de propósito

    problemas = validar_pos_geracao(linhas, caminho)
    assert any("Quantidade de linhas" in p for p in problemas)


def test_round_trip_detecta_valor_corrompido(tmp_path):
    df = pd.DataFrame([_linha_aprovada(VALOR=100.0)])
    linhas = formatar_linhas(df)
    caminho = tmp_path / "saida.txt"

    # escreve manualmente um TXT com valor diferente do esperado
    linha_corrompida = montar_linha_txt(linhas[0].debito, linhas[0].credito, "03/08/2026", "999.99", linhas[0].nova_descricao, 1)
    with open(caminho, "w", encoding="utf-8-sig", newline="") as f:
        f.write(linha_corrompida + "\r\n")

    problemas = validar_pos_geracao(linhas, caminho)
    assert any("VALOR lido" in p for p in problemas)


# ---------------------------------------------------------------------------
# Pré-condições
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not RELATORIO_VALIDACAO.exists(), reason="relatorio_validacao.xlsx não disponível")
def test_executar_bloqueia_se_status_geral_nao_aprovado(tmp_path):
    import shutil
    from openpyxl import load_workbook

    copia = tmp_path / "relatorio_bloqueado.xlsx"
    shutil.copy(RELATORIO_VALIDACAO, copia)
    wb = load_workbook(copia)
    ws = wb["RESUMO"]
    for row in ws.iter_rows():
        if row[0].value == "STATUS_GERAL":
            row[1].value = "BLOQUEADO_PARA_TXT"
    wb.save(copia)

    with pytest.raises(ExportacaoError):
        executar(str(copia), str(tmp_path / "saida.txt"), str(tmp_path / "auditoria.xlsx"))


# ---------------------------------------------------------------------------
# Ponta a ponta com o mês de exemplo
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not RELATORIO_VALIDACAO.exists(), reason="relatorio_validacao.xlsx não disponível")
def test_exportacao_completa_do_mes_de_exemplo(tmp_path):
    txt_saida = tmp_path / "saida.txt"
    auditoria = tmp_path / "auditoria.xlsx"

    status, linhas, problemas = executar(str(RELATORIO_VALIDACAO), str(txt_saida), str(auditoria))

    assert status == STATUS_TXT_VALIDADO
    assert problemas == []
    assert len(linhas) == 64
    assert txt_saida.exists()
    assert auditoria.exists()

    total = sum((l.valor_original for l in linhas), start=Decimal("0.00"))
    assert total == Decimal("99036.66")
