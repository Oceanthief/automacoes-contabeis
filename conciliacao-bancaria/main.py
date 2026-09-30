"""Ponto de entrada único do processamento mensal.

Uso:
    python main.py <EXTRATO> <CONTAS_PAGAS> <DE_PARA> <CORRECOES_MANUAIS> [DIRETORIO_SAIDA]

Este arquivo não decide nenhuma regra contábil — só chama
pipeline.processar_mes() e imprime o resultado de forma legível. Toda a
orquestração real está em pipeline.py; toda a lógica contábil está nos
módulos de cada fase (inalterados por este arquivo).
"""
from __future__ import annotations

import sys

from pipeline import STATUS_TXT_GERADO, processar_mes


def _imprimir_resultado(resultado) -> None:
    print("=" * 70)
    print(f"STATUS:      {resultado.status}")
    print(f"COMPETÊNCIA: {resultado.competencia}")
    print(f"SAÍDA:       {resultado.diretorio_saida}")
    print("=" * 70)

    print("\nFASES:")
    for fase in resultado.fases:
        marcador = {"OK": "✓", "FALHA": "✗", "BLOQUEADA": "■", "NAO_EXECUTADA": "-"}.get(fase.status, "?")
        print(f"  [{marcador}] {fase.nome}: {fase.status}")
        if fase.erro:
            print(f"      erro: {fase.erro}")

    if resultado.resumo:
        print("\nRESUMO:")
        for chave, valor in resultado.resumo.items():
            print(f"  {chave}: {valor}")

    if resultado.arquivos_gerados:
        print("\nARQUIVOS GERADOS:")
        for chave, caminho in resultado.arquivos_gerados.items():
            print(f"  {chave}: {caminho}")

    if resultado.pendencias:
        print("\nPENDÊNCIAS:")
        for p in resultado.pendencias:
            print(f"  - {p}")

    if resultado.erros:
        print("\nERROS:")
        for e in resultado.erros:
            print(f"  - {e}")

    print()


def main() -> int:
    if len(sys.argv) not in (5, 6):
        print(
            "Uso: python main.py <EXTRATO> <CONTAS_PAGAS> <DE_PARA> <CORRECOES_MANUAIS> [DIRETORIO_SAIDA]"
        )
        return 1

    extrato, contas_pagas, de_para, correcoes_manuais = sys.argv[1:5]
    diretorio_saida = sys.argv[5] if len(sys.argv) == 6 else "saida"

    resultado = processar_mes(extrato, contas_pagas, de_para, correcoes_manuais, diretorio_saida)
    _imprimir_resultado(resultado)

    return 0 if resultado.status == STATUS_TXT_GERADO else 1


if __name__ == "__main__":
    sys.exit(main())
