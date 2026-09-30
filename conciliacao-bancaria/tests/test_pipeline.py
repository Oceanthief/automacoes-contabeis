"""Testes de pipeline.py (orquestração) e main.py (CLI)."""
import filecmp
import sys
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from pipeline import STATUS_BLOQUEADO_PARA_TXT, STATUS_ERRO_PROCESSAMENTO, STATUS_TXT_GERADO, processar_mes

EXTRATO = RAIZ / "dados_exemplo" / "extrato_banco_2026_08.xlsx"
CONTAS_PAGAS = RAIZ / "dados_exemplo" / "contas_pagas_2026_08.xlsx"
DE_PARA = RAIZ / "dados_exemplo" / "de_para_contabil.xlsx"
CORRECOES_MANUAIS = RAIZ / "dados_exemplo" / "correcoes_manuais.xlsx"
TXT_REFERENCIA = RAIZ / "dados_exemplo" / "intermediarios" / "referencia_importacao_contabil.txt"


# ---------------------------------------------------------------------------
# Fluxo completo
# ---------------------------------------------------------------------------

def test_pipeline_completo_do_mes_de_exemplo(tmp_path):
    resultado = processar_mes(
        str(EXTRATO), str(CONTAS_PAGAS), str(DE_PARA), str(CORRECOES_MANUAIS), str(tmp_path / "saida")
    )

    assert resultado.status == STATUS_TXT_GERADO
    assert resultado.competencia == "2026-08"
    assert all(f.status == "OK" for f in resultado.fases)
    assert resultado.resumo["movimentacoes_bancarias"] == 58
    assert resultado.resumo["linhas_contabeis"] == 64
    assert resultado.resumo["linhas_aprovadas"] == 64
    assert resultado.resumo["linhas_bloqueadas"] == 0
    assert resultado.resumo["total_exportado"] == "99036.66"
    assert resultado.erros == []

    for chave in [
        "extrato_padronizado", "auditoria_sispag", "saida_enriquecida", "saida_enriquecida_sispag",
        "auditoria_de_para", "saida_contabil", "auditoria_granularizacao", "relatorio_validacao",
        "txt", "auditoria_exportacao_txt", "resumo_execucao",
    ]:
        assert Path(resultado.arquivos_gerados[chave]).exists(), f"{chave} não foi gerado"


def test_pipeline_organiza_pastas_por_competencia(tmp_path):
    resultado = processar_mes(
        str(EXTRATO), str(CONTAS_PAGAS), str(DE_PARA), str(CORRECOES_MANUAIS), str(tmp_path / "saida")
    )
    raiz = Path(resultado.diretorio_saida)
    assert raiz.name == "2026-08"
    for sub in ["intermediarios", "auditorias", "final", "logs"]:
        assert (raiz / sub).is_dir()
    assert (raiz / "logs" / "pipeline.log").exists()
    assert (raiz / "logs" / "resumo_execucao.json").exists()


# ---------------------------------------------------------------------------
# Interrupção em falha
# ---------------------------------------------------------------------------

def test_pipeline_falha_com_extrato_inexistente(tmp_path):
    resultado = processar_mes(
        "nao_existe.xlsx", str(CONTAS_PAGAS), str(DE_PARA), str(CORRECOES_MANUAIS), str(tmp_path / "saida")
    )
    assert resultado.status == STATUS_ERRO_PROCESSAMENTO
    assert len(resultado.erros) == 1
    assert "EXTRATO" in resultado.erros[0]


def test_pipeline_falha_no_de_para_interrompe_fases_seguintes(tmp_path):
    resultado = processar_mes(
        str(EXTRATO), str(CONTAS_PAGAS), "nao_existe.xlsx", str(CORRECOES_MANUAIS), str(tmp_path / "saida")
    )
    assert resultado.status == STATUS_ERRO_PROCESSAMENTO

    fases_por_nome = {f.nome: f.status for f in resultado.fases}
    # DE_PARA inexistente só é detectado na validação de entradas (fase 1) -
    # já que é um dos 4 arquivos obrigatórios
    assert fases_por_nome["1-validar-entradas"] == "FALHA"
    # nada além da fase 1 deveria ter sido registrado
    assert len(resultado.fases) == 1


# ---------------------------------------------------------------------------
# Nunca sobrescreve execução anterior
# ---------------------------------------------------------------------------

def test_pipeline_nao_sobrescreve_execucao_anterior_da_mesma_competencia(tmp_path):
    base = tmp_path / "saida"
    resultado1 = processar_mes(str(EXTRATO), str(CONTAS_PAGAS), str(DE_PARA), str(CORRECOES_MANUAIS), str(base))
    resultado2 = processar_mes(str(EXTRATO), str(CONTAS_PAGAS), str(DE_PARA), str(CORRECOES_MANUAIS), str(base))

    assert resultado1.diretorio_saida != resultado2.diretorio_saida
    assert Path(resultado1.diretorio_saida).exists()  # a primeira execução não foi apagada
    assert Path(resultado2.diretorio_saida).exists()
    assert any("já existia uma execução anterior" in p.lower() for p in resultado2.pendencias)


# ---------------------------------------------------------------------------
# Equivalência com a execução manual (fases isoladas)
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not TXT_REFERENCIA.exists(), reason="TXT de referência da execução manual não disponível")
def test_txt_do_pipeline_e_identico_ao_gerado_manualmente(tmp_path):
    resultado = processar_mes(
        str(EXTRATO), str(CONTAS_PAGAS), str(DE_PARA), str(CORRECOES_MANUAIS), str(tmp_path / "saida")
    )
    txt_pipeline = resultado.arquivos_gerados["txt"]
    assert filecmp.cmp(txt_pipeline, str(TXT_REFERENCIA), shallow=False)
