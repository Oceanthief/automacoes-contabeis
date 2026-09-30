"""
Geração da planilha contábil final (Fase 5).

Só gera arquivo quando TOTAL_PENDENCIAS = 0 (reaproveita o mesmo cálculo
de bloqueio da Fase 4 — nunca um cálculo "paralelo" que poderia divergir).
Gera dois arquivos separados:

  - `output/lancamentos_<empresa>_AAAA_MM.xlsx` — só as 6 colunas do
    layout oficial, sem NADA técnico a mais, pronta para virar TXT numa
    fase futura.
  - `output/auditoria_<empresa>_AAAA_MM.xlsx` — rastreabilidade completa
    até a linha original do extrato, mais a aba de reconciliação.

Fora de escopo: gerar o TXT de importação (isso é uma fase futura).
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "config"))
from leitura_extrato import Movimento
from classificador import Classificacao, STATUS_OK
from fluxo_mensal import processar_mes, ResultadoMes
import configuracoes as cfg

_FONTE = "Arial"

# ---------------------------------------------------------------------------
# Regra da sequência de lançamento (contrato de dados, Fase 1, Ambiguidade
# A3) — status explícito em código, não só em comentário: se algum dia
# essa regra precisar ser revista e marcada como pendente de novo, basta
# mudar `SEQUENCIA_REGRA_STATUS` em config/configuracoes.py para
# "PENDENTE_DEFINICAO" que a geração passa a bloquear sozinha, sem
# precisar lembrar de editar a lógica. Os nomes abaixo continuam
# existindo neste módulo (apontando para a config) só para não quebrar
# quem já importa `lancamentos_finais.SEQUENCIA_REGRA_STATUS`.
# ---------------------------------------------------------------------------
SEQUENCIA_REGRA_STATUS = cfg.SEQUENCIA_REGRA_STATUS
SEQUENCIA_REGRA_DESCRICAO = cfg.SEQUENCIA_REGRA_DESCRICAO

PADRAO_NOVO_HISTORICO = cfg.TEMPLATE_NOVO_HISTORICO  # configurável — ver montar_novo_historico()

_COLUNAS_OFICIAIS = ["DATA", "CONTA DEBITO", "CONTA CREDITO", "NOVO HISTORICO", "VALOR", "SEQUENCIA DE LANCAMENTO"]


class GeracaoBloqueadaError(Exception):
    """Levantada quando a Fase 5 não pode gerar o arquivo final — nunca
    silenciosamente; sempre com o motivo explícito."""


@dataclass
class LinhaFinal:
    source_row: int
    data: date
    conta_debito: str
    conta_credito: str
    novo_historico: str
    valor: float
    sequencia: int
    # campos de auditoria (não vão para o arquivo oficial):
    historico_original: str
    nome_identificado: Optional[str]
    regra_aplicada: str
    debito_origem: Optional[float]
    credito_origem: Optional[float]


# ---------------------------------------------------------------------------
# Validação de conta contábil
# ---------------------------------------------------------------------------

_PADRAO_DECIMAL_INTEIRO = re.compile(r"^\d+\.0+$")


def validar_conta(conta: Optional[str], contexto: str) -> str:
    """Levanta ValueError se a conta não passa nos critérios pedidos.
    Retorna a própria conta (como string) se estiver ok — nunca modifica
    o valor, só valida."""
    if conta is None:
        raise ValueError(f"{contexto}: conta ausente (None).")
    texto = str(conta).strip()
    if not texto:
        raise ValueError(f"{contexto}: conta vazia.")
    if texto.upper() == cfg.MARCADOR_PREENCHER_MANUAL:
        raise ValueError(f"{contexto}: conta ainda marcada como PREENCHER MANUAL — não pode chegar "
                          "na Fase 5 (isso indicaria que TOTAL_PENDENCIAS não era realmente 0).")
    if texto.lower() == "nan":
        raise ValueError(f"{contexto}: conta é 'NaN'.")
    if _PADRAO_DECIMAL_INTEIRO.match(texto):
        raise ValueError(f"{contexto}: conta '{texto}' parece um artefato decimal de número inteiro "
                          "(ex.: '11037.0') — contas devem ser texto puro.")
    return texto


def montar_novo_historico(chave_encontrada: Optional[str], historico_original: str) -> str:
    """Nome padronizado do DE-PARA -> 'VLR. REF {NOME}'. Quando não há
    nome do De-Para (ex.: item resolvido manualmente sem correspondência
    nenhuma), cai no histórico original como último recurso — nunca fica
    vazio, mas também nunca inventamos um nome que não veio de algum
    lugar real."""
    nome = chave_encontrada or historico_original
    # limpa anotações internas de auditoria que não devem vazar para o
    # histórico oficial (ex.: alias confirmado — ver classificador.py)
    nome = nome.split(" (alias confirmado de:")[0].strip()
    return PADRAO_NOVO_HISTORICO.format(nome=nome)


# ---------------------------------------------------------------------------
# Montagem das linhas finais
# ---------------------------------------------------------------------------

def montar_linhas_finais(
    classificacoes: list[Classificacao],
    movimentos_por_row: dict[int, Movimento],
) -> list[LinhaFinal]:
    """Constrói as linhas do lançamento final a partir das classificações
    já aprovadas (status OK). Assume que TOTAL_PENDENCIAS = 0 já foi
    conferido por quem chama — não refaz essa checagem aqui.
    """
    linhas: list[LinhaFinal] = []
    # ordem determinística: pela ordem cronológica do extrato (SOURCE_ROW)
    for c in sorted(classificacoes, key=lambda x: x.source_row):
        mov = movimentos_por_row.get(c.source_row)
        if mov is None or mov.data is None:
            raise GeracaoBloqueadaError(
                f"Movimento da linha {c.source_row} está aprovado (status OK) mas não tem data "
                "válida — isso não deveria acontecer se TOTAL_PENDENCIAS = 0 e indica uma "
                "inconsistência entre a Fase 4 e este módulo."
            )
        valor = mov.debito if mov.debito is not None else mov.credito
        if valor is None:
            raise GeracaoBloqueadaError(
                f"Movimento da linha {c.source_row} está aprovado mas não tem valor em DEBITO nem "
                "CREDITO — inconsistência entre a Fase 4 e este módulo."
            )
        linhas.append(LinhaFinal(
            source_row=c.source_row,
            data=mov.data,
            conta_debito=validar_conta(c.conta_debito, f"linha {c.source_row} (conta débito)"),
            conta_credito=validar_conta(c.conta_credito, f"linha {c.source_row} (conta crédito)"),
            novo_historico=montar_novo_historico(c.chave_encontrada, c.historico_original),
            valor=round(float(valor), 2),
            sequencia=0,  # preenchido depois, na ordem final
            historico_original=c.historico_original,
            nome_identificado=c.chave_encontrada,
            regra_aplicada=c.regra_aplicada,
            debito_origem=mov.debito,
            credito_origem=mov.credito,
        ))

    for i, linha in enumerate(linhas, start=1):
        linha.sequencia = i

    if not linhas:
        raise GeracaoBloqueadaError("Nenhum movimento aprovado neste mês — nada para gerar.")

    return linhas


# ---------------------------------------------------------------------------
# Reconciliação — roda ANTES de qualquer arquivo ser salvo
# ---------------------------------------------------------------------------

@dataclass
class RelatorioReconciliacao:
    ok: bool
    checagens: list[tuple[str, bool, str]]  # (nome, passou, detalhe)


def reconciliar(
    classificacoes_aprovadas: list[Classificacao],
    linhas: list[LinhaFinal],
) -> RelatorioReconciliacao:
    checagens: list[tuple[str, bool, str]] = []

    qtd_aprovados = len(classificacoes_aprovadas)
    qtd_linhas = len(linhas)
    checagens.append((
        "Quantidade de movimentos aprovados = quantidade de linhas finais",
        qtd_aprovados == qtd_linhas,
        f"{qtd_aprovados} aprovados vs. {qtd_linhas} linhas finais",
    ))

    soma_origem = round(sum(l.valor for l in linhas), 2)  # VALOR final já É o valor de origem (nunca alterado)
    soma_origem_bruta = round(sum(
        (l.debito_origem if l.debito_origem is not None else l.credito_origem) or 0.0 for l in linhas
    ), 2)
    checagens.append((
        "Soma dos VALORES de origem aprovados = soma dos VALORES finais",
        abs(soma_origem - soma_origem_bruta) < 0.005,
        f"origem={soma_origem_bruta} final={soma_origem}",
    ))

    source_rows = [l.source_row for l in linhas]
    duplicados = {r for r in source_rows if source_rows.count(r) > 1}
    checagens.append((
        "Nenhuma linha duplicada (por SOURCE_ROW)",
        len(duplicados) == 0,
        f"duplicados: {sorted(duplicados)}" if duplicados else "nenhuma",
    ))

    contas_vazias = [l.source_row for l in linhas if not l.conta_debito.strip() or not l.conta_credito.strip()]
    checagens.append((
        "Nenhuma conta vazia",
        len(contas_vazias) == 0,
        f"linhas com conta vazia: {contas_vazias}" if contas_vazias else "nenhuma",
    ))

    datas_vazias = [l.source_row for l in linhas if l.data is None]
    checagens.append((
        "Nenhuma data vazia",
        len(datas_vazias) == 0,
        f"linhas com data vazia: {datas_vazias}" if datas_vazias else "nenhuma",
    ))

    historicos_vazios = [l.source_row for l in linhas if not l.novo_historico.strip()]
    checagens.append((
        "Nenhum histórico vazio",
        len(historicos_vazios) == 0,
        f"linhas com histórico vazio: {historicos_vazios}" if historicos_vazios else "nenhuma",
    ))

    ok = all(passou for _, passou, _ in checagens)
    return RelatorioReconciliacao(ok=ok, checagens=checagens)


# ---------------------------------------------------------------------------
# Escrita dos dois arquivos
# ---------------------------------------------------------------------------

def _escrever_lancamentos_oficiais(linhas: list[LinhaFinal], caminho: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Lancamentos"

    fonte_cabecalho = Font(name=_FONTE, size=10, bold=True, color="FFFFFF")
    preenchimento_cabecalho = PatternFill("solid", fgColor="203864")
    fonte_normal = Font(name=_FONTE, size=10)

    for col, nome in enumerate(_COLUNAS_OFICIAIS, start=1):
        cel = ws.cell(row=1, column=col, value=nome)
        cel.font = fonte_cabecalho
        cel.fill = preenchimento_cabecalho
        cel.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"

    for i, linha in enumerate(linhas, start=2):
        ws.cell(row=i, column=1, value=linha.data).font = fonte_normal
        ws.cell(row=i, column=1).number_format = "DD/MM/YYYY"

        cel_debito = ws.cell(row=i, column=2, value=linha.conta_debito)
        cel_debito.font = fonte_normal
        cel_debito.number_format = "@"  # texto — nunca deixar o Excel reinterpretar como número

        cel_credito = ws.cell(row=i, column=3, value=linha.conta_credito)
        cel_credito.font = fonte_normal
        cel_credito.number_format = "@"

        ws.cell(row=i, column=4, value=linha.novo_historico).font = fonte_normal

        cel_valor = ws.cell(row=i, column=5, value=linha.valor)
        cel_valor.font = fonte_normal
        cel_valor.number_format = "#,##0.00"

        ws.cell(row=i, column=6, value=linha.sequencia).font = fonte_normal

    larguras = [14, 16, 16, 44, 14, 20]
    for col, largura in zip(range(1, 7), larguras):
        ws.column_dimensions[get_column_letter(col)].width = largura

    wb.save(caminho)


def _escrever_auditoria(
    linhas: list[LinhaFinal],
    relatorio: RelatorioReconciliacao,
    resultado_mes: ResultadoMes,
    caminho: Path,
) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Auditoria"

    colunas = [
        "SOURCE_ROW", "DATA", "HISTORICO_ORIGINAL", "DEBITO_ORIGEM", "CREDITO_ORIGEM",
        "NOME_IDENTIFICADO", "REGRA_APLICADA", "CONTA_DEBITO", "CONTA_CREDITO",
        "NOVO_HISTORICO", "VALOR", "SEQUENCIA_DE_LANCAMENTO",
    ]
    fonte_cabecalho = Font(name=_FONTE, size=10, bold=True, color="FFFFFF")
    preenchimento_cabecalho = PatternFill("solid", fgColor="375623")
    fonte_normal = Font(name=_FONTE, size=10)

    for col, nome in enumerate(colunas, start=1):
        cel = ws.cell(row=1, column=col, value=nome)
        cel.font = fonte_cabecalho
        cel.fill = preenchimento_cabecalho
    ws.freeze_panes = "A2"

    for i, l in enumerate(linhas, start=2):
        valores = [
            l.source_row, l.data, l.historico_original, l.debito_origem, l.credito_origem,
            l.nome_identificado, l.regra_aplicada, l.conta_debito, l.conta_credito,
            l.novo_historico, l.valor, l.sequencia,
        ]
        for col, v in enumerate(valores, start=1):
            cel = ws.cell(row=i, column=col, value=v)
            cel.font = fonte_normal
            if col == 2:
                cel.number_format = "DD/MM/YYYY"
            if col in (8, 9):
                cel.number_format = "@"

    larguras = [10, 12, 46, 12, 12, 30, 26, 14, 14, 40, 12, 18]
    for col, largura in zip(range(1, len(colunas) + 1), larguras):
        ws.column_dimensions[get_column_letter(col)].width = largura

    _escrever_aba_reconciliacao(wb, relatorio, resultado_mes)
    wb.save(caminho)


def _escrever_aba_reconciliacao(wb, relatorio: RelatorioReconciliacao, resultado_mes: ResultadoMes) -> None:
    ws = wb.create_sheet("Reconciliacao")
    fonte_normal = Font(name=_FONTE, size=10)
    fonte_titulo = Font(name=_FONTE, size=12, bold=True)
    fonte_ok = Font(name=_FONTE, size=10, color="006100")
    fonte_falha = Font(name=_FONTE, size=10, color="9C0006", bold=True)

    ws.cell(row=1, column=1, value=f"Reconciliação — {resultado_mes.mes:02d}/{resultado_mes.ano}").font = fonte_titulo
    ws.cell(row=2, column=1, value=f"Resultado geral: {'OK' if relatorio.ok else 'FALHOU'}").font = (
        fonte_ok if relatorio.ok else fonte_falha
    )

    linha = 4
    ws.cell(row=linha, column=1, value="Checagem").font = fonte_titulo
    ws.cell(row=linha, column=2, value="Resultado").font = fonte_titulo
    ws.cell(row=linha, column=3, value="Detalhe").font = fonte_titulo
    linha += 1
    for nome, passou, detalhe in relatorio.checagens:
        ws.cell(row=linha, column=1, value=nome).font = fonte_normal
        cel = ws.cell(row=linha, column=2, value="OK" if passou else "FALHOU")
        cel.font = fonte_ok if passou else fonte_falha
        ws.cell(row=linha, column=3, value=detalhe).font = fonte_normal
        linha += 1

    linha += 1
    ws.cell(row=linha, column=1, value="Regra da sequência de lançamento:").font = fonte_titulo
    linha += 1
    ws.cell(row=linha, column=1, value=f"Status: {SEQUENCIA_REGRA_STATUS}").font = fonte_normal
    linha += 1
    ws.cell(row=linha, column=1, value=SEQUENCIA_REGRA_DESCRICAO).font = fonte_normal

    for col, largura in zip("ABC", (55, 12, 70)):
        ws.column_dimensions[col].width = largura


# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------

@dataclass
class ResultadoGeracao:
    gerado: bool
    caminho_lancamentos: Optional[Path]
    caminho_auditoria: Optional[Path]
    linhas: list[LinhaFinal]
    relatorio_reconciliacao: Optional[RelatorioReconciliacao]
    motivo_bloqueio: Optional[str]


def gerar_planilha_final(
    caminho_extrato: str | Path,
    caminho_controle: str | Path,
    ano: int,
    mes: int,
    dir_saida: str | Path,
) -> ResultadoGeracao:
    """Ponto de entrada da Fase 5. Roda o fluxo mensal (Fases 2-4) para
    obter o estado ATUAL e aprovado dos movimentos, verifica
    TOTAL_PENDENCIAS = 0, e só então gera os dois arquivos. Nunca gera
    nada parcial: reconciliação roda antes de qualquer `wb.save`.
    """
    dir_saida = Path(dir_saida)

    resultado_mes = processar_mes(caminho_extrato, caminho_controle, ano, mes, dir_saida, silencioso=True)

    total_pendencias = sum(v for k, v in resultado_mes.contagem_status.items() if k != STATUS_OK)
    if total_pendencias > 0 or resultado_mes.bloqueado:
        motivo = (f"TOTAL_PENDENCIAS = {total_pendencias} (não é zero). Geração interrompida. "
                  f"Detalhe por status: {resultado_mes.contagem_status}. "
                  f"Resolva as pendências em {resultado_mes.caminho_pendencias.name} e rode de novo.")
        return ResultadoGeracao(False, None, None, [], None, motivo)

    if SEQUENCIA_REGRA_STATUS == "PENDENTE_DEFINICAO":
        motivo = ("A regra da SEQUENCIA DE LANCAMENTO está marcada como PENDENTE_DEFINICAO. "
                  "Geração bloqueada — não inventamos essa regra.")
        return ResultadoGeracao(False, None, None, [], None, motivo)

    classificacoes_aprovadas = [c for c in resultado_mes.classificacoes if c.status == STATUS_OK]
    linhas = montar_linhas_finais(classificacoes_aprovadas, resultado_mes.movimentos_por_row)

    relatorio = reconciliar(classificacoes_aprovadas, linhas)
    if not relatorio.ok:
        falhas = "; ".join(f"{n}: {d}" for n, passou, d in relatorio.checagens if not passou)
        motivo = f"Reconciliação falhou antes de salvar qualquer arquivo. Falhas: {falhas}"
        return ResultadoGeracao(False, None, None, linhas, relatorio, motivo)

    dir_saida.mkdir(parents=True, exist_ok=True)
    caminho_lancamentos = dir_saida / f"lancamentos_{cfg.PREFIXO_ARQUIVOS}_{ano:04d}_{mes:02d}.xlsx"
    caminho_auditoria = dir_saida / f"auditoria_{cfg.PREFIXO_ARQUIVOS}_{ano:04d}_{mes:02d}.xlsx"

    _escrever_lancamentos_oficiais(linhas, caminho_lancamentos)
    _escrever_auditoria(linhas, relatorio, resultado_mes, caminho_auditoria)

    return ResultadoGeracao(True, caminho_lancamentos, caminho_auditoria, linhas, relatorio, None)
