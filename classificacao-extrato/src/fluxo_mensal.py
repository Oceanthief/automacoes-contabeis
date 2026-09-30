"""
Orquestração do fluxo mensal (Fase 4).

Junta leitura (Fase 2), classificação (Fase 3) e pendências (Fase 4) num
único ponto de entrada, para poder ser chamado de novo depois da
conferência humana e reaproveitar os ajustes manuais salvos no arquivo
de pendências do mês.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from leitura_extrato import ler_extrato, Movimento
from classificador import carregar_depara, classificar_movimentos, Classificacao
from pendencias import (
    gerar_ou_atualizar_pendencias,
    escrever_aba_resolvidos,
    carregar_ajustes_manuais,
    carregar_source_rows_resolvidos,
    aplicar_ajustes_manuais,
    verificar_bloqueio,
    nome_arquivo_pendencias,
    MENSAGEM_BLOQUEIO,
)


@dataclass
class ResultadoMes:
    ano: int
    mes: int
    classificacoes: list[Classificacao]
    caminho_pendencias: Path
    bloqueado: bool
    contagem_status: dict[str, int]
    resolvidos_nesta_execucao: int
    movimentos_por_row: dict[int, Movimento]


def processar_mes(
    caminho_extrato: str | Path,
    caminho_controle: str | Path,
    ano: int,
    mes: int,
    dir_saida: str | Path,
    silencioso: bool = False,
) -> ResultadoMes:
    """Executa o fluxo completo de um mês:

      1. lê e normaliza o extrato (Fase 2) — nunca escreve no .xls;
      2. classifica automaticamente (Fase 3) — nunca inventa conta;
      3. reaproveita ajustes manuais já salvos no arquivo de pendências
         deste mês, se existirem — nunca reaplica regra automática sobre
         um valor que o conferente já decidiu;
      4. regrava o arquivo de pendências só com o que ainda está aberto,
         preservando preenchimentos parciais;
      5. registra numa aba separada o que foi resolvido nesta execução;
      6. avisa se o mês está liberado para finalização ou não.

    Pode ser chamado quantas vezes forem necessárias — é o mesmo fluxo
    para a primeira execução (sem pendências ainda) e para os
    reprocessamentos depois da conferência.
    """
    leitura = ler_extrato(caminho_extrato)
    depara = carregar_depara(caminho_controle)
    classificacoes_automaticas = classificar_movimentos(leitura.movimentos, depara)
    movimentos_por_row: dict[int, Movimento] = {m.source_row: m for m in leitura.movimentos}

    dir_saida = Path(dir_saida)
    caminho_pendencias = dir_saida / nome_arquivo_pendencias(ano, mes)

    ajustes = carregar_ajustes_manuais(caminho_pendencias)
    ja_resolvidos_anteriormente = carregar_source_rows_resolvidos(caminho_pendencias)
    resultado_ajuste = aplicar_ajustes_manuais(
        classificacoes_automaticas, movimentos_por_row, ajustes, ja_resolvidos_anteriormente
    )
    classificacoes_finais = resultado_ajuste.classificacoes

    # a data corrigida manualmente (quando houve) precisa valer também
    # para quem usar movimentos_por_row depois (ex.: a Fase 5, para
    # montar a coluna DATA do lançamento final) — nunca só como "senha"
    # para liberar a classificação.
    movimentos_por_row = {**movimentos_por_row, **resultado_ajuste.movimentos_corrigidos}

    caminho_pendencias = gerar_ou_atualizar_pendencias(
        classificacoes_finais, movimentos_por_row, depara, ano, mes, dir_saida
    )
    if resultado_ajuste.resolvidos_nesta_execucao:
        escrever_aba_resolvidos(caminho_pendencias, resultado_ajuste.resolvidos_nesta_execucao, ajustes)

    bloqueado, contagem = verificar_bloqueio(classificacoes_finais)

    if not silencioso:
        print(f"=== {mes:02d}/{ano} — {len(classificacoes_finais)} movimentos ===")
        for status in ("OK", "PENDENTE_MANUAL", "NAO_ENCONTRADO", "AMBIGUO", "REGRA_PENDENTE", "DADO_INVALIDO"):
            if contagem.get(status):
                print(f"  {status:16s} {contagem[status]}")
        if resultado_ajuste.resolvidos_nesta_execucao:
            print(f"  Resolvidos nesta execução via conferência manual: {len(resultado_ajuste.resolvidos_nesta_execucao)}")
        print(f"  Arquivo de pendências: {caminho_pendencias}")
        print(f"  {MENSAGEM_BLOQUEIO if bloqueado else 'LIBERADO PARA FINALIZAÇÃO — zero bloqueios.'}")
        print()

    return ResultadoMes(
        ano=ano, mes=mes, classificacoes=classificacoes_finais,
        caminho_pendencias=caminho_pendencias, bloqueado=bloqueado,
        contagem_status=contagem, resolvidos_nesta_execucao=len(resultado_ajuste.resolvidos_nesta_execucao),
        movimentos_por_row=movimentos_por_row,
    )
