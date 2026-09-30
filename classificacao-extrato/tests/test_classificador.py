"""Testes das regras do classificador (ordem de precedência e casos de borda)."""
import sys
from datetime import date
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(RAIZ / "src"), str(RAIZ / "config")]

import configuracoes as cfg  # noqa: E402
from classificador import (  # noqa: E402
    STATUS_AMBIGUO,
    STATUS_DADO_INVALIDO,
    STATUS_NAO_ENCONTRADO,
    STATUS_OK,
    STATUS_PENDENTE_MANUAL,
    STATUS_REGRA_PENDENTE,
    DeParaEntrada,
    carregar_depara,
    classificar_movimento,
)
from leitura_extrato import Movimento, _normalizar_texto  # noqa: E402

CONTROLE = RAIZ / "dados_exemplo" / "controle_fornecedores_clientes.xlsx"
EMPRESA = cfg.NOME_EMPRESA
BANCO = cfg.CONTA_SANTANDER


@pytest.fixture(scope="module")
def depara():
    return carregar_depara(CONTROLE)


def mov(historico, dia=10, credito=None, debito=None, flags=None):
    return Movimento(
        source_row=1, data=date(2026, 8, dia), data_bruta=None,
        historico_original=historico, historico_normalizado=_normalizar_texto(historico),
        debito=debito, credito=credito, flags=list(flags or []),
    )


# ------------------------------------------------------------------ transferências
def test_transferencia_recebida_da_filial_pr(depara):
    c = classificar_movimento(mov(f"TRANSFERENCIA DA CONTA SANTANDER {EMPRESA} - PR", credito=100), depara)
    assert (c.conta_debito, c.conta_credito, c.status) == (BANCO, cfg.CONTA_TRANSFERENCIA_PR, STATUS_OK)


def test_transferencia_enviada_para_filial_sc(depara):
    c = classificar_movimento(mov(f"TRANSFERENCIA PARA CONTA SANTANDER {EMPRESA} - SC", debito=100), depara)
    assert (c.conta_debito, c.conta_credito, c.status) == (cfg.CONTA_TRANSFERENCIA_SC, BANCO, STATUS_OK)


def test_saida_para_pr_em_credito_nao_e_presumida(depara):
    """Só o lado débito dessa frase foi confirmado: o crédito vira pendência."""
    c = classificar_movimento(mov(f"TRANSFERENCIA PARA CONTA SANTANDER {EMPRESA} - PR", credito=100), depara)
    assert c.status == STATUS_REGRA_PENDENTE
    assert c.conta_debito is None and c.conta_credito is None


# ------------------------------------------------------------------ salários
@pytest.mark.parametrize("dia, conta", [(3, cfg.CONTA_SALARIO_DIA_06), (6, cfg.CONTA_SALARIO_DIA_06),
                                        (9, cfg.CONTA_SALARIO_DIA_06), (17, cfg.CONTA_SALARIO_DIA_20),
                                        (20, cfg.CONTA_SALARIO_DIA_20), (23, cfg.CONTA_SALARIO_DIA_20)])
def test_salario_dentro_da_janela_de_3_dias(depara, dia, conta):
    c = classificar_movimento(mov(f"Pago: 1/08 - SALARIOS- {EMPRESA}", dia=dia, debito=500), depara)
    assert (c.conta_debito, c.conta_credito, c.status) == (conta, BANCO, STATUS_OK)


@pytest.mark.parametrize("dia", [1, 2, 10, 13, 16, 24, 31])
def test_salario_fora_da_janela_vira_pendencia(depara, dia):
    c = classificar_movimento(mov(f"Pago: 1/08 - SALARIOS- {EMPRESA}", dia=dia, debito=500), depara)
    assert c.status == STATUS_REGRA_PENDENTE


def test_pro_labore_vence_regra_de_salario(depara):
    historico = f"Pago: 1/08 - SALARIOS- {EMPRESA} - PRO LABORE SOCIO B"
    c = classificar_movimento(mov(historico, dia=6, debito=500), depara)
    assert c.regra_aplicada.startswith("pro_labore_prioritario")
    assert c.conta_debito == cfg.CONTA_PRO_LABORE


# ------------------------------------------------------------------ PAGO / RECEBIDO
def test_pago_usa_conta_do_controle_no_debito(depara):
    c = classificar_movimento(mov("Pago: 1/08 - DISTRIBUIDORA ALFA LTDA", debito=50), depara)
    assert (c.conta_debito, c.conta_credito, c.status) == ("21101", BANCO, STATUS_OK)


def test_recebido_usa_conta_do_controle_no_credito(depara):
    c = classificar_movimento(mov("Recebido: 9/08 - CLIENTE DELTA INDUSTRIA SA", credito=50), depara)
    assert (c.conta_debito, c.conta_credito, c.status) == (BANCO, "11302", STATUS_OK)


def test_nome_mais_longo_vence(depara):
    c = classificar_movimento(mov("Pago: 1/08 - BANCO SANTANDER - CONSORCIO", debito=50), depara)
    assert c.chave_encontrada == "BANCO SANTANDER - CONSORCIO"
    assert c.conta_debito == "13200"


def test_cadastro_duplicado_com_mesmas_contas_nao_e_ambiguo(depara):
    c = classificar_movimento(mov("Pago: 1/08 - BANCO SANTANDER", debito=50), depara)
    assert c.status == STATUS_OK


def test_empate_com_contas_diferentes_e_ambiguo():
    entradas = [
        DeParaEntrada("FORNECEDOR X", "FORNECEDOR X", "1", BANCO, 2),
        DeParaEntrada("FORNECEDOR Y", "FORNECEDOR Y", "2", BANCO, 3),
    ]
    c = classificar_movimento(mov("Pago: 1/08 - FORNECEDOR X FORNECEDOR Y", debito=50), entradas)
    assert c.status == STATUS_AMBIGUO
    assert c.conta_debito is None


def test_alias_confirmado_encontra_o_cadastro(depara):
    c = classificar_movimento(mov("Pago: 1/08 - J P DESPACHOS EXEMPLO", debito=50), depara)
    assert (c.conta_debito, c.status) == ("21103", STATUS_OK)


def test_preencher_manual_nunca_vira_conta(depara):
    c = classificar_movimento(mov("Pago: 1/08 - AP SERVICOS EXEMPLO", debito=50), depara)
    assert c.status == STATUS_PENDENTE_MANUAL
    assert c.conta_debito is None


def test_fornecedor_sem_cadastro(depara):
    c = classificar_movimento(mov("Pago: 1/08 - EMPRESA QUE NAO EXISTE", debito=50), depara)
    assert c.status == STATUS_NAO_ENCONTRADO


# ------------------------------------------------------------------ dado inválido
@pytest.mark.parametrize("flag", ["data_invalida", "sem_valor", "debito_e_credito_preenchidos",
                                  "historico_ausente"])
def test_flags_bloqueantes_da_leitura_viram_dado_invalido(depara, flag):
    c = classificar_movimento(mov("IOF", debito=1, flags=[flag]), depara)
    assert c.status == STATUS_DADO_INVALIDO
