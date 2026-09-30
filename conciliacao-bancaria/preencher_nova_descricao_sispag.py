"""Preenche NOVA DESCRICAO dos SISPAG FORNECEDORES usando auditoria_sispag.xlsx.

Escopo: só isso. Não refaz a conciliação, não mexe em DEBITO/CREDITO/DATA/
VALOR/SEQUENCIA/DESCRICAO de nenhuma linha, não altera nenhuma movimentação
que não seja SISPAG FORNECEDORES.

Uso:
    python preencher_nova_descricao_sispag.py <saida_enriquecida.xlsx> <auditoria_sispag.xlsx> <saida_enriquecida_sispag.xlsx>
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

from enriquecimento.merge_sispag import (
    DESCRICAO_SISPAG_FORNECEDORES,
    MergeSispagError,
    ResultadoMerge,
    processar,
)
from leitura.auditoria_sispag import AuditoriaSispagError, ler_auditoria_sispag
from leitura.extrato_padronizado import ExtratoPadronizadoError, ler_extrato_padronizado

COLUNAS_SAIDA = ["SEM INFORMACAO", "DEBITO", "CREDITO", "DATA", "VALOR", "SEM INFORMACAO", "DESCRICAO", "NOVA DESCRICAO", "SEQUENCIA"]


class PreenchimentoSispagError(Exception):
    pass


def _valor_ou_vazio(valor) -> str:
    """Preserva DEBITO/CREDITO exatamente como veio (string ou vazio), sem
    reinterpretar. Reaplica a mesma proteção contra perda de zero à esquerda
    da fase de enriquecimento."""
    if valor is None:
        return ""
    try:
        import pandas as pd

        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def aplicar(saida_df, resultados: list[ResultadoMerge]) -> list[dict]:
    """Reconstrói todas as 149 linhas, alterando NOVA DESCRICAO só onde o
    merge encontrou e validou um SISPAG preenchível. Todo o resto é cópia
    exata da linha original."""
    preenchimentos = {r.sequencia: r.nova_descricao for r in resultados if r.preenchido}

    linhas_saida = []
    for _, linha in saida_df.iterrows():
        sequencia = int(linha["SEQUENCIA"])
        nova_descricao = linha.get("NOVA DESCRICAO", "")
        nova_descricao = "" if (nova_descricao is None or _eh_nulo(nova_descricao)) else str(nova_descricao)

        if sequencia in preenchimentos:
            nova_descricao = preenchimentos[sequencia]

        linhas_saida.append(
            {
                "col_sem_info_1": "",
                "DEBITO": _valor_ou_vazio(linha["DEBITO"]),
                "CREDITO": _valor_ou_vazio(linha["CREDITO"]),
                "DATA": linha["DATA"],
                "VALOR": linha["VALOR"],
                "col_sem_info_2": "",
                "DESCRICAO": linha["DESCRICAO"],
                "NOVA DESCRICAO": nova_descricao,
                "SEQUENCIA": sequencia,
            }
        )
    return linhas_saida


def _eh_nulo(valor) -> bool:
    try:
        import pandas as pd

        return bool(pd.isna(valor))
    except (TypeError, ValueError):
        return False


def validar(saida_df, linhas_saida: list[dict]) -> None:
    """Garante que nada fora do escopo foi tocado."""
    if len(linhas_saida) != len(saida_df):
        raise PreenchimentoSispagError("Quantidade de linhas mudou — isso não deveria acontecer nesta fase.")

    for original, nova in zip(saida_df.iterrows(), linhas_saida):
        _, linha_original = original
        if str(linha_original["DESCRICAO"]).strip() == DESCRICAO_SISPAG_FORNECEDORES:
            continue  # só SISPAG pode ter NOVA DESCRICAO alterada
        original_nova_descricao = linha_original.get("NOVA DESCRICAO", "")
        original_nova_descricao = "" if _eh_nulo(original_nova_descricao) else str(original_nova_descricao)
        if nova["NOVA DESCRICAO"] != original_nova_descricao:
            raise PreenchimentoSispagError(
                f"SEQUENCIA {nova['SEQUENCIA']}: NOVA DESCRICAO de uma linha não-SISPAG foi alterada."
            )


def gravar_excel(linhas_saida: list[dict], caminho_saida: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Extrato Enriquecido (SISPAG)"

    fonte_cabecalho = Font(name="Arial", size=10, bold=True)
    fonte_dado = Font(name="Arial", size=10)

    ws.append(COLUNAS_SAIDA)
    for celula in ws[1]:
        celula.font = fonte_cabecalho

    for l in linhas_saida:
        ws.append(
            [
                l["col_sem_info_1"], l["DEBITO"], l["CREDITO"], l["DATA"], l["VALOR"],
                l["col_sem_info_2"], l["DESCRICAO"], l["NOVA DESCRICAO"], l["SEQUENCIA"],
            ]
        )

    for linha_celulas in ws.iter_rows(min_row=2):
        for celula in linha_celulas:
            celula.font = fonte_dado
        linha_celulas[1].number_format = "@"  # DEBITO como texto (zero à esquerda)
        linha_celulas[2].number_format = "@"  # CREDITO como texto
        linha_celulas[3].number_format = "DD/MM/YYYY"  # DATA
        linha_celulas[4].number_format = "#,##0.00"  # VALOR

    for coluna_celulas in ws.columns:
        largura = max((len(str(c.value)) for c in coluna_celulas if c.value is not None), default=0)
        ws.column_dimensions[coluna_celulas[0].column_letter].width = min(largura + 2, 80)

    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(caminho_saida)


def _imprimir_resumo(resultados: list[ResultadoMerge]) -> None:
    processados = len(resultados)
    preenchidos = [r for r in resultados if r.preenchido]
    ignorados_status = [r for r in resultados if r.motivo and r.motivo.startswith("STATUS não preenchível")]
    divergencia_data = [r for r in resultados if r.motivo and r.motivo.startswith("Divergência de DATA")]
    divergencia_valor = [r for r in resultados if r.motivo and r.motivo.startswith("Divergência de VALOR")]
    sem_correspondencia = [r for r in resultados if r.motivo == "SEQUENCIA sem correspondência na auditoria"]
    com_um_fornecedor = [r for r in preenchidos if "; " not in r.nova_descricao]
    com_multiplos_fornecedores = [r for r in preenchidos if "; " in r.nova_descricao]

    print(f"SISPAG processados:                {processados}")
    print(f"Preenchidos:                        {len(preenchidos)}")
    print(f"Ignorados por STATUS:                {len(ignorados_status)}")
    print(f"Divergências de DATA:                {len(divergencia_data)}")
    print(f"Divergências de VALOR:               {len(divergencia_valor)}")
    print(f"SEQUENCIAS sem correspondência:      {len(sem_correspondencia)}")
    print(f"SISPAG com 1 fornecedor:             {len(com_um_fornecedor)}")
    print(f"SISPAG com múltiplos fornecedores:   {len(com_multiplos_fornecedores)}")

    outros_nao_preenchidos = [
        r for r in resultados
        if not r.preenchido and r not in ignorados_status + divergencia_data + divergencia_valor + sem_correspondencia
    ]
    if outros_nao_preenchidos:
        print("\nOutros casos não preenchidos (revisão manual):")
        for r in outros_nao_preenchidos:
            print(f"  SEQ {r.sequencia}: {r.motivo}")

    if ignorados_status or divergencia_data or divergencia_valor or sem_correspondencia:
        print("\nCasos para revisão manual:")
        for r in ignorados_status + divergencia_data + divergencia_valor + sem_correspondencia:
            print(f"  SEQ {r.sequencia}: {r.motivo}")


def executar(caminho_saida_enriquecida: str, caminho_auditoria: str, caminho_saida: str) -> list[ResultadoMerge]:
    entrada = Path(caminho_saida_enriquecida)
    saida = Path(caminho_saida)

    if saida.resolve() == entrada.resolve():
        raise PreenchimentoSispagError("O arquivo de saída não pode ser igual ao arquivo de entrada.")

    saida_df = ler_extrato_padronizado(entrada)
    auditoria_df = ler_auditoria_sispag(caminho_auditoria)

    resultados = processar(saida_df, auditoria_df)
    linhas_saida = aplicar(saida_df, resultados)
    validar(saida_df, linhas_saida)
    gravar_excel(linhas_saida, saida)

    _imprimir_resumo(resultados)
    print(f"\nOK: gravado em {saida}")

    return resultados


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print(
            "Uso: python preencher_nova_descricao_sispag.py <saida_enriquecida.xlsx> "
            "<auditoria_sispag.xlsx> <saida_enriquecida_sispag.xlsx>"
        )
        sys.exit(1)
    try:
        executar(sys.argv[1], sys.argv[2], sys.argv[3])
    except (ExtratoPadronizadoError, AuditoriaSispagError, MergeSispagError, PreenchimentoSispagError) as exc:
        print(f"ERRO: {exc}")
        sys.exit(1)
