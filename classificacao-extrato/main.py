#!/usr/bin/env python3
"""
Classificação mensal do extrato — ponto de entrada único.

    python main.py preparar  --extrato input/extrato_MM_AA.xls
    python main.py finalizar --extrato input/extrato_MM_AA.xls

Ver README.md para o passo a passo completo, incluindo como resolver
pendências entre as duas etapas.

Este arquivo só ORQUESTRA os módulos já validados em src/ — nenhuma regra
contábil está aqui. As duas etapas (preparar/finalizar) chamam
exatamente as mesmas funções já testadas nas Fases 2 a 6.
"""

from __future__ import annotations

import argparse
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent / "config"))

import configuracoes as cfg  # noqa: E402
from leitura_extrato import ler_extrato  # noqa: E402
from fluxo_mensal import processar_mes  # noqa: E402
from lancamentos_finais import gerar_planilha_final  # noqa: E402
from exportar_txt import gerar_txt_importacao  # noqa: E402
from classificador import STATUS_OK  # noqa: E402
from log_execucao import registrar  # noqa: E402


def _inferir_ano_mes(movimentos) -> tuple[int, int]:
    """O mês/ano do arquivo é inferido pela data mais frequente entre os
    movimentos — assim o usuário não precisa digitar isso na linha de
    comando, só apontar o arquivo."""
    contagem = Counter((m.data.year, m.data.month) for m in movimentos if m.data is not None)
    if not contagem:
        print("ERRO: não foi possível determinar o mês/ano do extrato — nenhuma linha com "
              "data válida foi encontrada.")
        sys.exit(1)
    (ano, mes), _ = contagem.most_common(1)[0]
    return ano, mes


def _validar_arquivos_de_entrada(caminho_extrato: Path, caminho_controle: Path) -> None:
    if not caminho_extrato.exists():
        print(f"ERRO: extrato não encontrado: {caminho_extrato}")
        print(f"Coloque o arquivo do mês em {cfg.DIR_INPUT}/ e aponte --extrato para ele.")
        sys.exit(1)
    if not caminho_controle.exists():
        print(f"ERRO: Controle não encontrado: {caminho_controle}")
        print("O arquivo de controle de fornecedores e clientes precisa estar em "
              f"{cfg.DIR_INPUT}/ (ou aponte --controle para o caminho certo).")
        sys.exit(1)


# ---------------------------------------------------------------------------
# ETAPA 1 — PREPARAR
# ---------------------------------------------------------------------------

def cmd_preparar(args: argparse.Namespace) -> None:
    caminho_extrato = Path(args.extrato)
    caminho_controle = Path(args.controle) if args.controle else cfg.CAMINHO_CONTROLE_PADRAO
    dir_saida = Path(args.saida) if args.saida else cfg.DIR_OUTPUT

    _validar_arquivos_de_entrada(caminho_extrato, caminho_controle)

    leitura = ler_extrato(caminho_extrato)  # 1. ler
    ano, mes = _inferir_ano_mes(leitura.movimentos)
    print(f"Extrato: {caminho_extrato.name}  |  Mês identificado: {mes:02d}/{ano}")

    # 2. normalizar + 3. classificar + 4. validar + 5. gerar pendências
    # (tudo isso é o que processar_mes já faz, testado nas Fases 2-4)
    resultado = processar_mes(caminho_extrato, caminho_controle, ano, mes, dir_saida)

    qtd_ok = resultado.contagem_status.get(STATUS_OK, 0)
    qtd_pendencias = sum(v for k, v in resultado.contagem_status.items() if k != STATUS_OK)

    registrar(
        cfg.CAMINHO_LOG_EXECUCOES, "preparar", caminho_extrato,
        total_movimentos=len(resultado.classificacoes), qtd_ok=qtd_ok, qtd_pendencias=qtd_pendencias,
        arquivos_gerados=[resultado.caminho_pendencias],
    )

    print()
    if resultado.bloqueado:
        print(f"Existem {qtd_pendencias} pendência(s) — a preparação PAROU aqui, como esperado.")
        print(f"Abra: {resultado.caminho_pendencias}")
        print("Preencha CONTA_DEBITO_MANUAL e CONTA_CREDITO_MANUAL (e DATA_MANUAL quando pedido) "
              "na aba 'Pendencias', salve, e então rode:")
        print(f"  python main.py finalizar --extrato \"{caminho_extrato}\"")
        sys.exit(1)
    else:
        print("Nenhuma pendência. Pode seguir direto para:")
        print(f"  python main.py finalizar --extrato \"{caminho_extrato}\"")


# ---------------------------------------------------------------------------
# ETAPA 2 — FINALIZAR
# ---------------------------------------------------------------------------

def cmd_finalizar(args: argparse.Namespace) -> None:
    caminho_extrato = Path(args.extrato)
    caminho_controle = Path(args.controle) if args.controle else cfg.CAMINHO_CONTROLE_PADRAO
    dir_saida = Path(args.saida) if args.saida else cfg.DIR_OUTPUT

    _validar_arquivos_de_entrada(caminho_extrato, caminho_controle)

    leitura = ler_extrato(caminho_extrato)  # 1. reler os dados
    ano, mes = _inferir_ano_mes(leitura.movimentos)
    print(f"Extrato: {caminho_extrato.name}  |  Mês identificado: {mes:02d}/{ano}")

    # 2. incorporar ajustes autorizados + 3. validar zero pendências +
    # 4. gerar Excel final (tudo dentro de gerar_planilha_final, que por
    # sua vez chama processar_mes de novo — nunca confia num estado
    # antigo em memória)
    resultado_excel = gerar_planilha_final(caminho_extrato, caminho_controle, ano, mes, dir_saida)

    if not resultado_excel.gerado:
        print()
        print(f"NÃO FINALIZADO: {resultado_excel.motivo_bloqueio}")
        print("Rode 'python main.py preparar' de novo depois de resolver as pendências.")
        sys.exit(1)

    # 5. gerar TXT (inclui a validação de ida e volta própria da Fase 6)
    resultado_txt = gerar_txt_importacao(resultado_excel.caminho_lancamentos, ano, mes, dir_saida)

    if not resultado_txt.gerado:
        print()
        print(f"NÃO FINALIZADO (falha ao gerar o TXT): {resultado_txt.motivo_falha}")
        sys.exit(1)

    # 6. reconciliação final — resumo visível para quem está rodando
    soma_valores = round(sum(l.valor for l in resultado_excel.linhas), 2)
    arquivos_gerados = [resultado_excel.caminho_lancamentos, resultado_excel.caminho_auditoria,
                         resultado_txt.caminho_txt]

    registrar(
        cfg.CAMINHO_LOG_EXECUCOES, "finalizar", caminho_extrato,
        total_movimentos=len(resultado_excel.linhas), qtd_ok=len(resultado_excel.linhas), qtd_pendencias=0,
        arquivos_gerados=arquivos_gerados,
    )

    print()
    print("FINALIZADO COM SUCESSO.")
    print(f"  Registros aprovados = linhas da planilha final = linhas do TXT: {len(resultado_excel.linhas)}")
    print(f"  Total monetário: R$ {soma_valores:,.2f}".replace(",", "_").replace(".", ",").replace("_", "."))
    print(f"  Lançamentos: {resultado_excel.caminho_lancamentos}")
    print(f"  Auditoria:   {resultado_excel.caminho_auditoria}")
    print(f"  TXT:         {resultado_txt.caminho_txt}")
    print("  Nenhuma importação foi feita no ERP — isso continua manual, fora deste programa.")


# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        prog="python main.py",
        description="Fluxo mensal de classificação do extrato (preparar -> conferir -> finalizar).",
    )
    subparsers = parser.add_subparsers(dest="comando", required=True)

    p_preparar = subparsers.add_parser("preparar", help="Lê, classifica e gera as pendências do mês.")
    p_preparar.add_argument("--extrato", required=True, help="Caminho do extrato .xls do mês (em input/).")
    p_preparar.add_argument("--controle", required=False, help="Caminho do Controle (padrão: input/Controle...xlsx).")
    p_preparar.add_argument("--saida", required=False, help="Diretório de saída (padrão: output/).")
    p_preparar.set_defaults(func=cmd_preparar)

    p_finalizar = subparsers.add_parser("finalizar", help="Gera a planilha final e o TXT, se não houver pendências.")
    p_finalizar.add_argument("--extrato", required=True, help="Caminho do extrato .xls do mês (em input/).")
    p_finalizar.add_argument("--controle", required=False, help="Caminho do Controle (padrão: input/Controle...xlsx).")
    p_finalizar.add_argument("--saida", required=False, help="Diretório de saída (padrão: output/).")
    p_finalizar.set_defaults(func=cmd_finalizar)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
