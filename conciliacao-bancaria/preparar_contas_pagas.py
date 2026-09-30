"""Prepara e normaliza o arquivo de Contas Pagas em uma coleção de pagamentos únicos.

Escopo: só a camada de dados (leitura → normalização → reconstrução por
TITULO+DATA_EFETIVA_CH). Não faz conciliação com o Extrato/SISPAG — isso é a
próxima fase.

Uso:
    python preparar_contas_pagas.py <contas_pagas.xlsx>
"""
from __future__ import annotations

import sys
from pathlib import Path

from leitura.contas_pagas import (
    CONTA_ALVO,
    ContasPagasError,
    ler_contas_pagas,
    listar_divergencias,
    montar_pagamentos,
)


def executar(caminho: str) -> None:
    df = ler_contas_pagas(caminho)
    candidatos, fora_do_escopo = montar_pagamentos(df)
    divergentes = listar_divergencias(candidatos + fora_do_escopo)

    print(f"Linhas de origem:              {len(df)}")
    print(f"Pagamentos únicos (candidatos, {Path(caminho).name} / {CONTA_ALVO}): {len(candidatos)}")
    print(f"Pagamentos fora do escopo (outras contas): {len(fora_do_escopo)}")
    print(f"Divergências VALOR x soma(VALOR_CLASS) > R$ 0,01: {len(divergentes)}")

    if divergentes:
        print("\nDetalhe das divergências:")
        for p in divergentes:
            print(
                f"  TITULO {p.titulo} ({p.data_efetiva}) — {p.nome} — "
                f"VALOR={p.valor_original}  soma(VALOR_CLASS)={p.valor_classificado}  "
                f"diferença={p.divergencia}"
            )

    if fora_do_escopo:
        contas = sorted({p.conta for p in fora_do_escopo})
        print(f"\nContas fora do escopo encontradas: {contas}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python preparar_contas_pagas.py <contas_pagas.xlsx>")
        sys.exit(1)
    try:
        executar(sys.argv[1])
    except ContasPagasError as exc:
        print(f"ERRO: {exc}")
        sys.exit(1)
