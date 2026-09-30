#!/usr/bin/env python3
"""
Ponto único de entrada da automação CONTAMAX — uso mensal.

Uso:
    python main.py "extrato_do_mes.xls"

O que este script faz (e nada além disso — não recalcula nada, não muda
nenhuma regra contábil, só organiza a chamada do que já foi construído e
auditado nos Prompts 2 a 5):

    1. valida o arquivo de entrada
    2. valida a estrutura do extrato (cabeçalho, datas)
    3. detecta o período de referência (AAAA-MM)
    4. processa as movimentações (motor contábil, não alterado)
    5. gera o Excel de conferência
    6. confere o Excel gerado (reabre o arquivo, checa abas e cabeçalhos)
    7. gera o TXT de importação do ERP
    8. confere o TXT gerado (já validado internamente por gerar_txt_contamax)
    9. apresenta um resumo em português simples

Os arquivos saem em:
    output/AAAA-MM/contamax_AAAA-MM.xlsx
    output/AAAA-MM/contamax_AAAA-MM.txt

Se esses arquivos já existirem, o script para e avisa — nunca sobrescreve
silenciosamente.
"""
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

from contamax.exportador_txt import EstrategiaSequencia, gerar_txt_contamax
from contamax.exportador_xlsx import gerar_excel_conferencia
from contamax.leitura_xls import CabecalhoNaoEncontradoError, DataInvalidaError
from contamax.processamento import processar_extrato_xls
from contamax.resumo import montar_resumo
from contamax.validacoes import ErroCriticoValidacao
from contamax.validacoes_txt import ErroValidacaoTxt

EXTENSOES_SUPORTADAS = (".xls",)
DIRETORIO_SAIDA_BASE = Path("output")
CAMINHO_ESPECIFICACAO = Path(__file__).resolve().parent / "docs" / "especificacao_contamax.md"


class ErroAmigavel(Exception):
    """
    Um erro com mensagem já pronta para uma pessoa não técnica ler — sem
    jargão de programação, sem stack trace. main() só imprime `str(erro)`
    e para; não faz nenhuma tradução adicional.
    """


# ---------------------------------------------------------------------------
# Passo 1-2: validações antes de processar qualquer coisa
# ---------------------------------------------------------------------------
def validar_arquivo_de_entrada(caminho: str) -> Path:
    caminho_path = Path(caminho)

    if not caminho_path.exists():
        raise ErroAmigavel(
            f"Não encontrei o arquivo '{caminho}'.\n"
            f"Confira se o caminho está certo e se o arquivo está na pasta que você indicou."
        )

    if not caminho_path.is_file():
        raise ErroAmigavel(f"'{caminho}' é uma pasta, não um arquivo. Aponte para o arquivo .xls do extrato.")

    if caminho_path.suffix.lower() not in EXTENSOES_SUPORTADAS:
        raise ErroAmigavel(
            f"O arquivo '{caminho_path.name}' tem a extensão '{caminho_path.suffix}', que esta automação "
            f"ainda não sabe ler. Por enquanto ela só entende arquivos '.xls' (o formato antigo do Excel, "
            f"que é o que o internet banking do Santander exporta). Se o seu arquivo é '.xlsx', abra-o no "
            f"Excel e salve novamente escolhendo 'Pasta de Trabalho do Excel 97-2003 (*.xls)'."
        )

    return caminho_path


def verificar_sem_decisoes_pendentes() -> None:
    """
    Checagem rápida: se o arquivo de especificação ainda contiver o texto
    'PENDENTE DE DEFINIÇÃO', significa que alguma decisão de negócio (ver
    docs/especificacao_contamax.md) nunca foi respondida — e essa automação
    não deveria gerar nada em cima de uma regra que ninguém confirmou.
    Se o arquivo de especificação não estiver presente, a checagem é
    simplesmente pulada (não bloqueia o uso mensal por falta de um arquivo
    de documentação).
    """
    if not CAMINHO_ESPECIFICACAO.exists():
        return
    conteudo = CAMINHO_ESPECIFICACAO.read_text(encoding="utf-8")
    if "PENDENTE DE DEFINIÇÃO" in conteudo:
        raise ErroAmigavel(
            "A especificação da automação (docs/especificacao_contamax.md) ainda tem uma decisão "
            "marcada como PENDENTE DE DEFINIÇÃO. Antes de gerar arquivos para importar no ERP, essa "
            "decisão precisa ser respondida — gerar em cima de uma regra não confirmada arrisca criar "
            "lançamentos contábeis errados."
        )


# ---------------------------------------------------------------------------
# Passo 3: caminhos de saída (e proteção contra sobrescrita silenciosa)
# ---------------------------------------------------------------------------
def preparar_caminhos_saida(periodo: str) -> tuple[Path, Path]:
    diretorio = DIRETORIO_SAIDA_BASE / periodo
    caminho_xlsx = diretorio / f"contamax_{periodo}.xlsx"
    caminho_txt = diretorio / f"contamax_{periodo}.txt"

    existentes = [str(c) for c in (caminho_xlsx, caminho_txt) if c.exists()]
    if existentes:
        lista = "\n".join(f"  - {c}" for c in existentes)
        raise ErroAmigavel(
            f"Já existe(m) resultado(s) gerado(s) para o período {periodo}:\n{lista}\n\n"
            f"Para não sobrescrever por engano, apague ou renomeie esse(s) arquivo(s) antes de rodar de "
            f"novo — ou confirme se você já não processou esse mês hoje mais cedo."
        )

    diretorio.mkdir(parents=True, exist_ok=True)
    return caminho_xlsx, caminho_txt


# ---------------------------------------------------------------------------
# Passo 6: conferência do Excel recém-gerado (reabre o arquivo do zero)
# ---------------------------------------------------------------------------
def conferir_excel_gerado(caminho_xlsx: Path) -> None:
    wb = openpyxl.load_workbook(str(caminho_xlsx))
    if wb.sheetnames != ["LANCAMENTOS", "AUDITORIA_DIARIA"]:
        raise ErroAmigavel(
            f"O Excel gerado em '{caminho_xlsx}' não tem as abas esperadas (LANCAMENTOS e "
            f"AUDITORIA_DIARIA). Isso indica um problema na geração — não use este arquivo."
        )
    cabecalho_lancamentos = [c.value for c in wb["LANCAMENTOS"][1]]
    if cabecalho_lancamentos != ["DATA", "CONTA DEBITO", "CONTA CREDITO", "HISTORICO", "VALOR"]:
        raise ErroAmigavel(
            f"O Excel gerado em '{caminho_xlsx}' tem cabeçalhos inesperados na aba LANCAMENTOS. "
            f"Isso indica um problema na geração — não use este arquivo."
        )


# ---------------------------------------------------------------------------
# Fluxo principal
# ---------------------------------------------------------------------------
def executar(caminho_entrada: str) -> int:
    print("[1/9] Verificando arquivo de entrada...")
    caminho_xls = validar_arquivo_de_entrada(caminho_entrada)

    print("[2/9] Verificando decisões pendentes na especificação...")
    verificar_sem_decisoes_pendentes()

    print("[3/9] Lendo e validando a estrutura do extrato (cabeçalho, datas)...")
    try:
        resultados = processar_extrato_xls(str(caminho_xls))
    except CabecalhoNaoEncontradoError:
        raise ErroAmigavel(
            f"Não encontrei as colunas esperadas (Data, Histórico, Crédito, Débito, Saldo) no arquivo "
            f"'{caminho_xls.name}'. Confira se este é mesmo um extrato bancário no formato de sempre "
            f"(pode ter sido exportado de forma diferente, ou ser outro tipo de relatório)."
        )
    except DataInvalidaError as exc:
        raise ErroAmigavel(
            f"Encontrei uma data que não consigo entender em '{caminho_xls.name}'.\n"
            f"Detalhe: {exc}\n"
            f"Abra o arquivo no Excel e confira se essa linha está certa."
        )

    if not resultados:
        raise ErroAmigavel(
            f"O extrato '{caminho_xls.name}' não tem nenhuma movimentação para processar. "
            f"Confira se é o arquivo certo."
        )

    resumo_periodo = montar_resumo(resultados)
    periodo = resumo_periodo.periodo
    print(f"      Período detectado: {periodo}")

    print("[4/9] Processando movimentações (agrupamento por data + regra contábil)...")
    print(f"      {resumo_periodo.dias_processados} dia(s) processado(s).")

    caminho_xlsx, caminho_txt = preparar_caminhos_saida(periodo)

    print("[5/9] Gerando Excel de conferência...")
    try:
        gerar_excel_conferencia(resultados, str(caminho_xlsx))
    except ErroCriticoValidacao as exc:
        raise ErroAmigavel(
            "Encontrei problema(s) que impedem gerar o Excel com segurança:\n"
            + "\n".join(f"  - {p}" for p in exc.problemas)
        )

    print("[6/9] Conferindo o Excel gerado...")
    conferir_excel_gerado(caminho_xlsx)
    print(f"      OK: {caminho_xlsx}")

    print("[7/9] Gerando TXT de importação do ERP...")
    estrategia_sequencia = EstrategiaSequencia()  # decisão confirmada: especificação, seção 5.2
    try:
        resumo_txt = gerar_txt_contamax(resultados, str(caminho_txt), estrategia_sequencia)
    except ErroValidacaoTxt as exc:
        raise ErroAmigavel(
            "O TXT gerado falhou na conferência e foi removido. Problema(s) encontrado(s):\n"
            + "\n".join(f"  - {p}" for p in exc.problemas)
        )

    print("[8/9] TXT já conferido automaticamente (0 erros estruturais).")
    print(f"      OK: {caminho_txt}")

    print("[9/9] Resumo")
    print()
    print(f"Período: {periodo}")
    print(f"Dias processados: {resumo_periodo.dias_processados}")
    print(f"Total de resgates: {resumo_periodo.total_resgates}")
    print(f"Total de aplicações: {resumo_periodo.total_aplicacoes}")
    print(f"Pendências (empate ou valor inválido — revisar manualmente): {resumo_periodo.pendencias}")
    print()
    print(f"Excel de conferência: {caminho_xlsx}")
    print(f"TXT para importar no ERP: {caminho_txt} ({resumo_txt.lancamentos} lançamento(s))")

    return 0


def main(argv: list[str]) -> int:
    if len(argv) != 1:
        print('Uso: python main.py "caminho/para/extrato.xls"')
        return 2

    try:
        return executar(argv[0])
    except ErroAmigavel as erro:
        print()
        print("=" * 60)
        print("NÃO FOI POSSÍVEL CONCLUIR")
        print("=" * 60)
        print(str(erro))
        return 1
    except Exception as erro:  # última rede de proteção — nunca mostrar um stack trace cru
        print()
        print("=" * 60)
        print("ALGO INESPERADO ACONTECEU")
        print("=" * 60)
        print(
            "Isso não é um erro esperado do processo normal. Anote a mensagem abaixo e "
            "mostre para quem cuida desta automação:"
        )
        print(f"  ({type(erro).__name__}) {erro}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
