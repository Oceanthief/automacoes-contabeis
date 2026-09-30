"""
Leitura do extrato bancário .xls e normalização para Movimentacao.

Esta é a única camada do pacote que conhece o xlrd — todo o resto do
pipeline (contamax.regras, contamax.processamento) trabalha só com os tipos
de contamax.models. Isolar a dependência de biblioteca de planilha aqui
facilita trocar de biblioteca no futuro (ex.: se um mês vier em .xlsx em vez
de .xls) sem tocar na regra de negócio.

Importante (conforme a especificação, seção 1): o cabeçalho é localizado
procurando pelos nomes de coluna, não por posição fixa (linha 1) — isso é o
que garante que o código não dependa do texto do título (que muda a cada
mês/conta) nem da suposição de que o layout de linhas é sempre idêntico.
"""
from __future__ import annotations

import xlrd

from .models import LinhaComErro, Movimentacao
from .valores import (
    CTYPE_BLANK,
    CTYPE_EMPTY,
    CTYPE_TEXT,
    ValorInvalidoError,
    extrair_valor_monetario,
)

COLUNAS_ESPERADAS = ("Data", "Histórico", "Crédito", "Débito", "Saldo")
LINHAS_TOTALIZACAO = {"SALDO ANTERIOR", "SALDO FINAL"}
LINHAS_MAX_BUSCA_CABECALHO = 10  # não precisa varrer a planilha inteira


class CabecalhoNaoEncontradoError(RuntimeError):
    """Nenhuma das primeiras linhas contém o cabeçalho esperado."""


class DataInvalidaError(ValueError):
    """
    Célula de Data não é um valor de data válido.

    Ao contrário de um valor monetário inválido (que afeta só um dia e é
    coletado em LinhaComErro), uma data inválida impede até de saber a qual
    dia a linha pertence — não há como agrupar algo sem chave de agrupamento.
    Por isso este erro interrompe a leitura imediatamente (falha controlada,
    em vez de silenciosamente ignorar a linha ou adivinhar uma data).
    """


def _localizar_colunas(sheet: xlrd.sheet.Sheet) -> tuple[int, dict[str, int]]:
    """
    Procura, nas primeiras linhas da planilha, a linha de cabeçalho real
    (a que contém os nomes das colunas) e devolve (índice_da_linha, mapa
    nome_coluna -> índice_da_coluna).
    """
    limite = min(LINHAS_MAX_BUSCA_CABECALHO, sheet.nrows)
    for linha_idx in range(limite):
        valores_linha = [
            str(sheet.cell(linha_idx, col).value).strip() for col in range(sheet.ncols)
        ]
        mapa = {nome: valores_linha.index(nome) for nome in COLUNAS_ESPERADAS if nome in valores_linha}
        if len(mapa) == len(COLUNAS_ESPERADAS):
            return linha_idx, mapa
    raise CabecalhoNaoEncontradoError(
        f"Não encontrei uma linha com as colunas {COLUNAS_ESPERADAS} "
        f"nas primeiras {limite} linhas da planilha."
    )


def _celula_vazia(sheet: xlrd.sheet.Sheet, linha: int, col: int) -> bool:
    cell = sheet.cell(linha, col)
    return cell.ctype in (CTYPE_EMPTY, CTYPE_BLANK) or (
        cell.ctype == CTYPE_TEXT and cell.value == ""
    )


def _e_linha_de_totalizacao(sheet: xlrd.sheet.Sheet, linha: int, col_historico: int, col_debito: int, col_credito: int) -> bool:
    """
    Detecta linhas como 'SALDO ANTERIOR' / 'SALDO FINAL' por dois critérios
    combinados (conforme risco #2 registrado na especificação):
    1. o texto do histórico bate com um dos rótulos conhecidos, OU
    2. estruturalmente, tanto Débito quanto Crédito estão vazios nessa linha
       (nenhuma movimentação real deixaria as duas colunas em branco).
    O critério (2) é o que protege contra o banco mudar o texto do rótulo
    em outro mês.
    """
    historico = str(sheet.cell(linha, col_historico).value).strip().upper()
    if historico in LINHAS_TOTALIZACAO:
        return True
    return _celula_vazia(sheet, linha, col_debito) and _celula_vazia(sheet, linha, col_credito)


def ler_extrato(caminho: str) -> tuple[list[Movimentacao], list[LinhaComErro]]:
    """
    Lê o arquivo .xls e retorna (movimentacoes_validas, linhas_com_erro).

    Linhas de totalização (SALDO ANTERIOR/FINAL) são descartadas silenciosamente
    (não são erro, são esperadas). Linhas com valor monetário inválido viram
    LinhaComErro (o dia correspondente é marcado como erro em
    contamax.processamento, sem interromper a leitura das outras linhas).
    Uma data inválida interrompe a leitura inteira (ver DataInvalidaError).
    """
    workbook = xlrd.open_workbook(caminho)
    sheet = workbook.sheet_by_index(0)

    linha_cabecalho, colunas = _localizar_colunas(sheet)
    col_data = colunas["Data"]
    col_historico = colunas["Histórico"]
    col_credito = colunas["Crédito"]
    col_debito = colunas["Débito"]

    movimentacoes: list[Movimentacao] = []
    linhas_com_erro: list[LinhaComErro] = []

    for linha in range(linha_cabecalho + 1, sheet.nrows):
        if _e_linha_de_totalizacao(sheet, linha, col_historico, col_debito, col_credito):
            continue

        cell_data = sheet.cell(linha, col_data)
        if cell_data.ctype != xlrd.XL_CELL_DATE:
            raise DataInvalidaError(
                f"Linha {linha + 1}: célula de Data não é uma data válida "
                f"(ctype={cell_data.ctype}, valor={cell_data.value!r})."
            )
        data_convertida = xlrd.xldate_as_datetime(cell_data.value, workbook.datemode).date()

        historico = str(sheet.cell(linha, col_historico).value).strip()

        try:
            debito = extrair_valor_monetario(
                sheet.cell(linha, col_debito).ctype, sheet.cell(linha, col_debito).value
            )
            credito = extrair_valor_monetario(
                sheet.cell(linha, col_credito).ctype, sheet.cell(linha, col_credito).value
            )
        except ValorInvalidoError as exc:
            linhas_com_erro.append(
                LinhaComErro(numero_linha=linha + 1, data=data_convertida, mensagem=str(exc))
            )
            continue

        movimentacoes.append(
            Movimentacao(data=data_convertida, debito=debito, credito=credito, historico=historico)
        )

    return movimentacoes, linhas_com_erro
