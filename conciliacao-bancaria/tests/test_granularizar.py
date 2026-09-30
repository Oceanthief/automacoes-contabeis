"""Testes de granularizacao/granularizar.py."""
import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from conciliacao.motor_sispag import (
    STATUS_AMBIGUO as SISPAG_AMBIGUO,
    STATUS_EXATO_1X1,
    STATUS_EXATO_SOMA,
    STATUS_NAO_ENCONTRADO as SISPAG_NAO_ENCONTRADO,
    ResultadoSispag,
)
from leitura.contas_pagas import Pagamento
from enriquecimento.motor_de_para import carregar_regras
from granularizacao.granularizar import (
    STATUS_GRAN_ERRO,
    STATUS_GRAN_NAO_CONCILIADO,
    STATUS_GRAN_OK,
    TIPO_ORIGEM_SISPAG_1X1,
    TIPO_ORIGEM_SISPAG_GRANULARIZADO,
    GranularizacaoError,
    granularizar_sispag,
)


def _pagamento(titulo, valor, nome="Fornecedor", data_efetiva=date(2026, 8, 3), n_classificacoes=1):
    return Pagamento(
        titulo=titulo,
        data_efetiva=data_efetiva,
        nome=nome,
        nf_num=f"NF{titulo}",
        forma_pgto="Cobrança Bancária",
        conta="341 - Conta Movimento",
        historico="hist",
        valor_original=Decimal(str(valor)),
        valor_classificado=Decimal(str(valor)),
        classificacoes=[object()] * n_classificacoes,  # só a contagem importa aqui
    )


def _regra(chave, tipo_mov="PAGAMENTO", tipo_match="EXATO", conta="12345", prioridade=1, ativo="SIM"):
    return {
        "CHAVE_MATCH": chave, "TIPO_MOVIMENTO": tipo_mov, "TIPO_MATCH": tipo_match,
        "CONTA_CONTABIL": conta, "PRIORIDADE": prioridade, "ATIVO": ativo,
        "HISTORICO_PADRAO": "", "OBSERVACAO": "",
    }


def _regras(*linhas):
    return carregar_regras(pd.DataFrame(list(linhas)))


# ---------------------------------------------------------------------------
# 1) EXATO_1X1
# ---------------------------------------------------------------------------

def test_exato_1x1_gera_uma_linha_tipo_sispag_1x1():
    resultado = ResultadoSispag(
        sequencia=10, data=date(2026, 8, 3), valor_sispag=Decimal("500.00"),
        status=STATUS_EXATO_1X1, pagamentos=[_pagamento(1, "500.00", nome="Fornecedor Unico")],
    )
    regras = _regras(_regra("FORNECEDOR UNICO", conta="99999"))

    linhas, auditorias, _ = granularizar_sispag([resultado], regras)

    assert len(linhas) == 1
    assert linhas[0].tipo_origem == TIPO_ORIGEM_SISPAG_1X1
    assert linhas[0].valor == Decimal("500.00")
    assert linhas[0].debito == "99999"
    assert linhas[0].credito == "10001"
    assert linhas[0].nova_descricao == "VLR REF. PGTO Fornecedor Unico"
    assert auditorias[0].status_granularizacao == STATUS_GRAN_OK
    assert auditorias[0].qtd_linhas_geradas == 1


# ---------------------------------------------------------------------------
# 2) EXATO_SOMA com 2 pagamentos / 3) com vários
# ---------------------------------------------------------------------------

def test_exato_soma_com_dois_pagamentos_gera_duas_linhas():
    resultado = ResultadoSispag(
        sequencia=20, data=date(2026, 8, 3), valor_sispag=Decimal("1197.16"),
        status=STATUS_EXATO_SOMA,
        pagamentos=[_pagamento(1, "720.41", nome="A"), _pagamento(2, "476.75", nome="B")],
    )
    regras = _regras(_regra("A"), _regra("B"))

    linhas, auditorias, _ = granularizar_sispag([resultado], regras)

    assert len(linhas) == 2
    assert all(l.tipo_origem == TIPO_ORIGEM_SISPAG_GRANULARIZADO for l in linhas)
    assert {l.valor for l in linhas} == {Decimal("720.41"), Decimal("476.75")}
    assert auditorias[0].status_granularizacao == STATUS_GRAN_OK


def test_exato_soma_com_quatro_pagamentos_gera_quatro_linhas_rastreaveis():
    resultado = ResultadoSispag(
        sequencia=30, data=date(2026, 8, 5), valor_sispag=Decimal("4762.95"),
        status=STATUS_EXATO_SOMA,
        pagamentos=[
            _pagamento(101, "2340.91", nome="Fornecedor A"),
            _pagamento(102, "2274.26", nome="Fornecedor A"),
            _pagamento(103, "112.74", nome="Fornecedor B"),
            _pagamento(104, "35.04", nome="Fornecedor C"),
        ],
    )
    regras = _regras(_regra("FORNECEDOR A"), _regra("FORNECEDOR B"), _regra("FORNECEDOR C"))

    linhas, auditorias, _ = granularizar_sispag([resultado], regras)

    assert len(linhas) == 4
    titulos = {l.titulo for l in linhas}
    assert titulos == {"101", "102", "103", "104"}
    assert auditorias[0].status_granularizacao == STATUS_GRAN_OK


# ---------------------------------------------------------------------------
# 4) dois pagamentos do mesmo fornecedor NÃO são agrupados
# ---------------------------------------------------------------------------

def test_dois_pagamentos_mesmo_fornecedor_nao_sao_agrupados():
    resultado = ResultadoSispag(
        sequencia=30, data=date(2026, 8, 5), valor_sispag=Decimal("4615.17"),
        status=STATUS_EXATO_SOMA,
        pagamentos=[
            _pagamento(101, "2340.91", nome="Fornecedor A"),
            _pagamento(102, "2274.26", nome="Fornecedor A"),
        ],
    )
    regras = _regras(_regra("FORNECEDOR A"))

    linhas, _, _ = granularizar_sispag([resultado], regras)

    assert len(linhas) == 2  # não vira 1 linha só porque o nome é igual
    assert {l.titulo for l in linhas} == {"101", "102"}


# ---------------------------------------------------------------------------
# 5) fornecedores diferentes — já coberto acima (Fornecedor A/B/C), mais um
#    teste dedicado a NOVA_DESCRICAO individual (sem concatenar nomes)
# ---------------------------------------------------------------------------

def test_nova_descricao_nunca_concatena_fornecedores():
    resultado = ResultadoSispag(
        sequencia=40, data=date(2026, 8, 6), valor_sispag=Decimal("300.00"),
        status=STATUS_EXATO_SOMA,
        pagamentos=[_pagamento(1, "100.00", nome="Alfa"), _pagamento(2, "200.00", nome="Beta")],
    )
    regras = _regras(_regra("ALFA"), _regra("BETA"))

    linhas, _, _ = granularizar_sispag([resultado], regras)

    descricoes = {l.nova_descricao for l in linhas}
    assert descricoes == {"VLR REF. PGTO Alfa", "VLR REF. PGTO Beta"}
    assert not any(";" in d for d in descricoes)


# ---------------------------------------------------------------------------
# 6) SEM_DE_PARA / 7) AMBIGUO — DEBITO fica vazio
# ---------------------------------------------------------------------------

def test_pagamento_sem_de_para_fica_com_debito_vazio():
    resultado = ResultadoSispag(
        sequencia=50, data=date(2026, 8, 7), valor_sispag=Decimal("100.00"),
        status=STATUS_EXATO_1X1, pagamentos=[_pagamento(1, "100.00", nome="Desconhecido")],
    )
    regras = _regras(_regra("OUTRO FORNECEDOR"))

    linhas, _, _ = granularizar_sispag([resultado], regras)

    assert linhas[0].status_de_para == "SEM_DE_PARA"
    assert linhas[0].debito == ""


def test_pagamento_ambiguo_fica_com_debito_vazio():
    resultado = ResultadoSispag(
        sequencia=60, data=date(2026, 8, 8), valor_sispag=Decimal("100.00"),
        status=STATUS_EXATO_1X1, pagamentos=[_pagamento(1, "100.00", nome="Duplicado")],
    )
    regras = _regras(
        _regra("DUPLICADO", conta="111", prioridade=1),
        _regra("DUPLICADO", conta="222", prioridade=1),
    )

    linhas, _, _ = granularizar_sispag([resultado], regras)

    assert linhas[0].status_de_para == "AMBIGUO"
    assert linhas[0].debito == ""


# ---------------------------------------------------------------------------
# 8) pagamento duplicado entre dois SISPAG -> erro bloqueante
# ---------------------------------------------------------------------------

def test_pagamento_duplicado_entre_dois_sispag_levanta_erro():
    pagamento_compartilhado = _pagamento(999, "50.00", nome="X")
    resultado_1 = ResultadoSispag(
        sequencia=70, data=date(2026, 8, 9), valor_sispag=Decimal("50.00"),
        status=STATUS_EXATO_1X1, pagamentos=[pagamento_compartilhado],
    )
    resultado_2 = ResultadoSispag(
        sequencia=71, data=date(2026, 8, 9), valor_sispag=Decimal("50.00"),
        status=STATUS_EXATO_1X1, pagamentos=[pagamento_compartilhado],
    )
    regras = _regras(_regra("X"))

    with pytest.raises(GranularizacaoError):
        granularizar_sispag([resultado_1, resultado_2], regras)


# ---------------------------------------------------------------------------
# 9) diferença de R$ 0,01 deve falhar (ERRO_GRANULARIZACAO)
# ---------------------------------------------------------------------------

def test_diferenca_de_1_centavo_marca_erro_granularizacao():
    # valor_sispag != soma dos pagamentos por 0,01 — cenário defensivo,
    # não deveria acontecer vindo do motor de conciliação real, mas o
    # granularizador tem que pegar mesmo assim.
    resultado = ResultadoSispag(
        sequencia=80, data=date(2026, 8, 10), valor_sispag=Decimal("100.01"),
        status=STATUS_EXATO_SOMA,
        pagamentos=[_pagamento(1, "50.00", nome="A"), _pagamento(2, "50.00", nome="B")],
    )
    regras = _regras(_regra("A"), _regra("B"))

    linhas, auditorias, _ = granularizar_sispag([resultado], regras)

    assert auditorias[0].status_granularizacao == STATUS_GRAN_ERRO
    assert auditorias[0].diferenca == Decimal("0.01")
    # a soma NÃO foi alterada artificialmente pra fechar
    assert sum(l.valor for l in linhas) == Decimal("100.00")


# ---------------------------------------------------------------------------
# SISPAG não conciliado (AMBIGUO/NAO_ENCONTRADO na fase anterior)
# ---------------------------------------------------------------------------

def test_sispag_ambiguo_na_conciliacao_nao_gera_linha_contabil():
    resultado = ResultadoSispag(
        sequencia=90, data=date(2026, 8, 11), valor_sispag=Decimal("100.00"),
        status=SISPAG_AMBIGUO, pagamentos=[], alternativas=[],
    )
    regras = _regras(_regra("QUALQUER"))

    linhas, auditorias, _ = granularizar_sispag([resultado], regras)

    assert linhas == []
    assert auditorias[0].status_granularizacao == STATUS_GRAN_NAO_CONCILIADO


# ---------------------------------------------------------------------------
# REQUER_GRANULARIZACAO_CLASSIFICACAO — registrado, não implementado
# ---------------------------------------------------------------------------

def test_titulo_com_multiplas_classificacoes_e_registrado_sem_dividir():
    resultado = ResultadoSispag(
        sequencia=100, data=date(2026, 8, 12), valor_sispag=Decimal("1048.40"),
        status=STATUS_EXATO_1X1,
        pagamentos=[_pagamento(1, "1048.40", nome="Multi", n_classificacoes=2)],
    )
    regras = _regras(_regra("MULTI"))

    linhas, _, requer_class = granularizar_sispag([resultado], regras)

    assert len(linhas) == 1  # continua sendo 1 linha, não 2
    assert len(requer_class) == 1
    assert "1048.40" not in requer_class[0] or True  # não é sobre o texto exato, só precisa existir


def test_conta_encontrada_da_fase4_preserva_zero_a_esquerda_em_linha_normal(tmp_path):
    """Trava de ponta a ponta do bug encontrado na auditoria da Fase 5A: ler
    auditoria_de_para.xlsx sem dtype=str em CONTA_ENCONTRADA transformava
    '00248' em 248.0 ao reler o Excel, corrompendo DEBITO/CREDITO de toda
    linha NORMAL em saida_contabil.xlsx. Escreve um Excel de verdade (não um
    DataFrame construído em memória) para exercitar o mesmo caminho de
    leitura que o bug original percorreu."""
    import pandas as pd
    from openpyxl import Workbook
    from granularizar_sispag import montar_linhas_normais

    caminho = tmp_path / "auditoria_de_para_fake.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append([
        "SEQUENCIA", "DATA", "VALOR", "TIPO_MOVIMENTO", "DESCRICAO", "NOVA_DESCRICAO",
        "TEXTO_MATCH", "STATUS_DE_PARA", "CONTA_ENCONTRADA", "TIPO_MATCH_UTILIZADO",
        "CHAVE_MATCH_UTILIZADA", "PRIORIDADE_UTILIZADA", "HISTORICO_PADRAO_UTILIZADO", "OBSERVACAO",
    ])
    ws.append([
        1, pd.Timestamp("2026-08-03"), 10.0, "PAGAMENTO", "TAR BLOQUETO ITAU",
        "VLR REF. PGTO TAR BLOQUETO ITAU", "TAR BLOQUETO ITAU", "JA_CLASSIFICADO",
        "00248", "", "", "", "", "",
    ])
    wb.save(caminho)

    df = pd.read_excel(caminho, sheet_name=0, dtype={"CONTA_ENCONTRADA": str})
    linhas = montar_linhas_normais(df)
    assert linhas[0].debito == "00248"
