"""Fase 7B: exportador TXT.

Fonte de verdade: especificacao_txt.md (Fase 7A). Este módulo NÃO decide
nada de contabilidade — só pega as linhas já aprovadas pela Fase 6 e
transforma em texto no formato exato já comprovado: 8 campos entre aspas
duplas, separados por vírgula, CRLF, UTF-8 com BOM, sem cabeçalho, sem
rodapé, campos 1 e 6 sempre vazios.

Fluxo: receber lançamentos aprovados -> formatar -> escrever -> reler ->
validar -> comparar (round-trip) -> auditar. Nenhuma etapa altera dado
contábil; qualquer divergência vira erro, nunca correção silenciosa.
"""
from __future__ import annotations

import sys
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

from leitura.relatorio_validacao import (
    STATUS_GERAL_APROVADO,
    RelatorioValidacaoError,
    ler_linhas_aprovadas,
    ler_status_geral,
)
from normalizacao import normalizar_conta_contabil, valor_para_decimal

ENCODING = "utf-8-sig"  # UTF-8 com BOM — confirmado na Fase 7A
QUEBRA_DE_LINHA = "\r\n"
QTD_CAMPOS = 8

STATUS_TXT_VALIDADO = "TXT_GERADO_VALIDADO"
STATUS_TXT_COM_ERRO = "TXT_GERADO_COM_ERRO"

COLUNAS_AUDITORIA = [
    "SEQUENCIA_CONTABIL", "SEQUENCIA_BANCO", "DATA", "DEBITO", "CREDITO", "VALOR",
    "NOVA_DESCRICAO", "LINHA_TXT", "STATUS_EXPORTACAO", "OBSERVACAO",
]


class ExportacaoError(Exception):
    """Erro bloqueante: pré-condição não satisfeita ou pós-validação/round-trip falhou."""


# ---------------------------------------------------------------------------
# Formatação — cada função implementa exatamente uma regra de especificacao_txt.md
# ---------------------------------------------------------------------------

def _formatar_data(valor) -> str:
    """DD/MM/AAAA — única máscara observada na especificação."""
    ts = pd.Timestamp(valor)
    if pd.isna(ts):
        raise ExportacaoError(f"DATA inválida para exportação: {valor!r}")
    return ts.strftime("%d/%m/%Y")


def _formatar_valor(valor) -> str:
    """Ponto decimal, 2 casas, sempre positivo, nunca vírgula/separador de milhar."""
    decimal_valor = valor_para_decimal(valor)
    if decimal_valor <= 0:
        raise ExportacaoError(f"VALOR deveria ser positivo (Fase 6 já deveria ter bloqueado): {decimal_valor}")
    return f"{decimal_valor:.2f}"


def _formatar_conta(valor) -> str:
    """Preserva o código exatamente como veio da saída validada — nunca
    força um tamanho fixo (ex.: o código '8685', confirmado pelo usuário
    como uma exceção legítima de 4 dígitos, não pode ser alterado aqui)."""
    texto = normalizar_conta_contabil(valor)
    if not texto:
        raise ExportacaoError(f"Conta vazia não pode ser exportada (a Fase 6 já deveria ter bloqueado): {valor!r}")
    return texto


def _escapar_campo(valor: str) -> str:
    """Escape padrão de CSV: aspas internas viram aspas duplicadas. Nenhuma
    linha da especificação exigiu isso até agora, mas é mecânica de CSV, não
    uma regra de negócio inventada."""
    return str(valor).replace('"', '""')


def montar_linha_txt(debito: str, credito: str, data: str, valor: str, nova_descricao: str, sequencia) -> str:
    """8 campos entre aspas duplas, separados por vírgula — campos 1 e 6 vazios."""
    campos = ["", debito, credito, data, valor, "", nova_descricao, str(sequencia)]
    return ",".join(f'"{_escapar_campo(c)}"' for c in campos)


@dataclass
class LinhaExportada:
    sequencia_contabil: int
    sequencia_banco: int
    data_original: object
    debito: str
    credito: str
    valor_original: Decimal
    nova_descricao: str
    linha_txt: str


def formatar_linhas(linhas_aprovadas: pd.DataFrame) -> list[LinhaExportada]:
    resultado = []
    for _, linha in linhas_aprovadas.iterrows():
        debito = _formatar_conta(linha["DEBITO"])
        credito = _formatar_conta(linha["CREDITO"])
        data_fmt = _formatar_data(linha["DATA"])
        valor_fmt = _formatar_valor(linha["VALOR"])
        nova_descricao = "" if pd.isna(linha["NOVA_DESCRICAO"]) else str(linha["NOVA_DESCRICAO"])
        sequencia = int(linha["SEQUENCIA_CONTABIL"])

        linha_txt = montar_linha_txt(debito, credito, data_fmt, valor_fmt, nova_descricao, sequencia)

        resultado.append(
            LinhaExportada(
                sequencia_contabil=sequencia,
                sequencia_banco=int(linha["SEQUENCIA_BANCO"]),
                data_original=linha["DATA"],
                debito=debito,
                credito=credito,
                valor_original=valor_para_decimal(linha["VALOR"]),
                nova_descricao=nova_descricao,
                linha_txt=linha_txt,
            )
        )
    return resultado


def escrever_txt(linhas: list[LinhaExportada], caminho_saida: Path) -> None:
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    with open(caminho_saida, "w", encoding=ENCODING, newline="") as f:
        for linha in linhas:
            f.write(linha.linha_txt + QUEBRA_DE_LINHA)


# ---------------------------------------------------------------------------
# Pós-validação: reler o TXT gerado do zero, sem confiar no que acabou de
# ser escrito em memória
# ---------------------------------------------------------------------------

def reler_e_reconstruir(caminho_txt: Path) -> list[dict]:
    """Lê o TXT gerado byte a byte com as regras da especificação (não reusa
    nenhum objeto Python já em memória) e devolve os campos reconstruídos."""
    with open(caminho_txt, "rb") as f:
        conteudo = f.read()

    if not conteudo.startswith(b"\xef\xbb\xbf"):
        raise ExportacaoError("TXT gerado não começa com BOM UTF-8 — encoding incorreto.")

    texto = conteudo.decode("utf-8-sig")

    if not texto.endswith("\r\n"):
        raise ExportacaoError("TXT gerado não termina com CRLF — quebra de linha incorreta na última linha.")

    linhas_texto = texto.split("\r\n")
    linhas_texto = [l for l in linhas_texto if l != ""]  # descarta só o "" residual do split final

    reconstruidas = []
    for numero, linha_txt in enumerate(linhas_texto, start=1):
        if "\n" in linha_txt or "\r" in linha_txt:
            raise ExportacaoError(f"Linha {numero}: quebra de linha inesperada dentro do conteúdo.")
        campos = linha_txt.split('","')
        if len(campos) != QTD_CAMPOS:
            raise ExportacaoError(
                f"Linha {numero}: {len(campos)} campos encontrados, esperado {QTD_CAMPOS}. Conteúdo: {linha_txt!r}"
            )
        campos[0] = campos[0].lstrip('"')
        campos[-1] = campos[-1].rstrip('"')

        reconstruidas.append(
            {
                "linha_numero": numero,
                "campo1_vazio": campos[0],
                "debito": campos[1],
                "credito": campos[2],
                "data": campos[3],
                "valor": campos[4],
                "campo6_vazio": campos[5],
                "nova_descricao": campos[6],
                "sequencia": campos[7],
                "linha_txt": linha_txt,
            }
        )
    return reconstruidas


def validar_pos_geracao(linhas_esperadas: list[LinhaExportada], caminho_txt: Path) -> list[str]:
    """Executa as 10 checagens de pós-validação + round-trip. Devolve lista
    de problemas encontrados (vazia = tudo certo)."""
    problemas: list[str] = []

    reconstruidas = reler_e_reconstruir(caminho_txt)  # já valida encoding e CRLF, levanta se falhar

    # 2) quantidade de linhas
    if len(reconstruidas) != len(linhas_esperadas):
        problemas.append(
            f"Quantidade de linhas divergente: TXT tem {len(reconstruidas)}, esperado {len(linhas_esperadas)}."
        )
        return problemas  # sem contagem batendo, não faz sentido comparar campo a campo

    # 3-7) round-trip campo a campo
    for esperada, lida in zip(linhas_esperadas, reconstruidas):
        if lida["campo1_vazio"] != "" or lida["campo6_vazio"] != "":
            problemas.append(f"SEQUENCIA_CONTABIL {esperada.sequencia_contabil}: campo vazio fixo não está vazio.")
        if lida["debito"] != esperada.debito:
            problemas.append(
                f"SEQUENCIA_CONTABIL {esperada.sequencia_contabil}: DEBITO lido '{lida['debito']}' != esperado '{esperada.debito}'."
            )
        if lida["credito"] != esperada.credito:
            problemas.append(
                f"SEQUENCIA_CONTABIL {esperada.sequencia_contabil}: CREDITO lido '{lida['credito']}' != esperado '{esperada.credito}'."
            )
        data_esperada = _formatar_data(esperada.data_original)
        if lida["data"] != data_esperada:
            problemas.append(
                f"SEQUENCIA_CONTABIL {esperada.sequencia_contabil}: DATA lida '{lida['data']}' != esperada '{data_esperada}'."
            )
        valor_lido = valor_para_decimal(lida["valor"])
        if valor_lido != esperada.valor_original:
            problemas.append(
                f"SEQUENCIA_CONTABIL {esperada.sequencia_contabil}: VALOR lido {valor_lido} != esperado {esperada.valor_original}."
            )
        if lida["nova_descricao"].replace('""', '"') != esperada.nova_descricao:
            problemas.append(f"SEQUENCIA_CONTABIL {esperada.sequencia_contabil}: NOVA_DESCRICAO divergente após releitura.")
        if int(lida["sequencia"]) != esperada.sequencia_contabil:
            problemas.append(
                f"SEQUENCIA_CONTABIL {esperada.sequencia_contabil}: SEQUENCIA lida '{lida['sequencia']}' diverge."
            )

    # 10) totais
    total_esperado = sum((l.valor_original for l in linhas_esperadas), start=Decimal("0.00"))
    total_lido = sum((valor_para_decimal(r["valor"]) for r in reconstruidas), start=Decimal("0.00"))
    if total_esperado != total_lido:
        problemas.append(f"Total monetário divergente: TXT tem {total_lido}, esperado {total_esperado}.")

    # 8-9) cabeçalho/rodapé: não deve existir nenhum
    if reconstruidas and reconstruidas[0]["sequencia"] != str(linhas_esperadas[0].sequencia_contabil):
        problemas.append("Primeira linha do TXT não é a primeira SEQUENCIA_CONTABIL esperada — possível cabeçalho indevido.")

    return problemas


# ---------------------------------------------------------------------------
# Auditoria
# ---------------------------------------------------------------------------

def gravar_auditoria(
    linhas: list[LinhaExportada], problemas: list[str], status: str, caminho_saida: Path
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Auditoria Exportacao TXT"

    fonte_cabecalho = Font(name="Arial", size=10, bold=True)
    fonte_dado = Font(name="Arial", size=10)

    ws.append(COLUNAS_AUDITORIA)
    for c in ws[1]:
        c.font = fonte_cabecalho

    status_linha = "EXPORTADO_OK" if not problemas else "EXPORTADO_COM_DIVERGENCIA_GLOBAL"
    observacao_linha = "" if not problemas else "Ver problemas gerais no console/relatório de execução."

    for l in linhas:
        ws.append(
            [
                l.sequencia_contabil, l.sequencia_banco, l.data_original, l.debito, l.credito,
                float(l.valor_original), l.nova_descricao, l.linha_txt, status_linha, observacao_linha,
            ]
        )

    for linha_celulas in ws.iter_rows(min_row=2):
        for c in linha_celulas:
            c.font = fonte_dado
        linha_celulas[2].number_format = "DD/MM/YYYY"
        linha_celulas[3].number_format = "@"
        linha_celulas[4].number_format = "@"
        linha_celulas[5].number_format = "#,##0.00"

    for coluna_celulas in ws.columns:
        largura = max((len(str(c.value)) for c in coluna_celulas if c.value is not None), default=0)
        ws.column_dimensions[coluna_celulas[0].column_letter].width = min(largura + 2, 80)

    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(caminho_saida)


# ---------------------------------------------------------------------------
# Orquestração
# ---------------------------------------------------------------------------

def executar(caminho_relatorio_validacao: str, caminho_txt_saida: str, caminho_auditoria: str):
    status_geral = ler_status_geral(caminho_relatorio_validacao)
    if status_geral != STATUS_GERAL_APROVADO:
        raise ExportacaoError(
            f"Pré-condição não satisfeita: STATUS_GERAL='{status_geral}', esperado '{STATUS_GERAL_APROVADO}'. "
            "TXT não será gerado."
        )

    linhas_aprovadas_df = ler_linhas_aprovadas(caminho_relatorio_validacao)
    if len(linhas_aprovadas_df) == 0:
        raise ExportacaoError("Nenhuma linha APROVADA encontrada — nada para exportar.")

    linhas = formatar_linhas(linhas_aprovadas_df)

    caminho_txt = Path(caminho_txt_saida)
    escrever_txt(linhas, caminho_txt)

    problemas = validar_pos_geracao(linhas, caminho_txt)
    status_final = STATUS_TXT_COM_ERRO if problemas else STATUS_TXT_VALIDADO

    gravar_auditoria(linhas, problemas, status_final, Path(caminho_auditoria))

    total = sum((l.valor_original for l in linhas), start=Decimal("0.00"))
    print(f"Lançamentos exportados: {len(linhas)}")
    print(f"Total monetário exportado: R$ {total}")
    print(f"Pós-validação/round-trip: {'SEM PROBLEMAS' if not problemas else f'{len(problemas)} PROBLEMA(S)'}")
    for p in problemas:
        print(f"  - {p}")
    print(f"STATUS: {status_final}")
    print(f"\nTXT gravado em: {caminho_txt}")
    print(f"Auditoria gravada em: {caminho_auditoria}")

    return status_final, linhas, problemas


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print(
            "Uso: python exportacao_txt.py <relatorio_validacao.xlsx> <saida.txt> <auditoria_exportacao_txt.xlsx>"
        )
        sys.exit(1)
    try:
        status_final, _, problemas = executar(sys.argv[1], sys.argv[2], sys.argv[3])
        sys.exit(0 if status_final == STATUS_TXT_VALIDADO else 1)
    except (RelatorioValidacaoError, ExportacaoError) as exc:
        print(f"ERRO: {exc}")
        sys.exit(1)
