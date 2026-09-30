"""Pipeline mensal: encadeia as fases já implementadas e auditadas.

Este módulo NÃO decide nenhuma regra contábil — cada fase continua sendo
responsabilidade exclusiva do seu próprio módulo (leitura/, enriquecimento/,
conciliacao/, granularizacao/, validacao/, exportacao_txt.py). Aqui só
existe: em que ordem chamar, onde salvar cada saída, o que fazer quando uma
fase falha ou bloqueia a seguinte.

Fluxo (a numeração é a mesma da tarefa que originou este orquestrador):
  1. validar arquivos de entrada
  2. processar Extrato               -> fase1_padronizar_extrato
  3. processar Contas Pagas          -> leitura.contas_pagas
  4. conciliar SISPAG                -> conciliar_sispag
  5. criar NOVA DESCRICAO            -> enriquecer_extrato + preencher_nova_descricao_sispag
  6. aplicar regras fixas (tarifas)  -> já dentro de enriquecer_extrato (mesma passada, ver nota)
  7. executar DE-PARA                -> executar_de_para
  8. granularizar SISPAG             -> granularizar_sispag
  9. executar validação contábil     -> validar_saida_contabil
 10. gerar TXT, só se aprovado       -> exportacao_txt
 11. gerar resumo final              -> ResultadoPipeline + log

Nota sobre os passos 5 e 6: enriquecer_extrato.py decide, linha a linha, se é
uma tarifa (regra fixa) ou uma descrição comum (NOVA DESCRICAO) numa única
passada — são a mesma função de negócio desde a Fase 3, então o pipeline os
representa como uma fase só («5-6») em vez de forçar uma separação artificial
que exigiria reescrever aquele módulo.
"""
from __future__ import annotations

import json
import shutil
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Optional

from conciliacao.motor_sispag import (
    STATUS_AMBIGUO as SISPAG_AMBIGUO,
    STATUS_EXATO_1X1,
    STATUS_EXATO_SOMA,
    STATUS_NAO_ENCONTRADO as SISPAG_NAO_ENCONTRADO,
)
from conciliar_sispag import executar as conciliar_sispag_executar
from enriquecer_extrato import executar as enriquecer_extrato_executar
from executar_de_para import executar as executar_de_para_executar
from exportacao_txt import STATUS_GERAL_APROVADO, STATUS_TXT_VALIDADO, executar as exportacao_txt_executar
from fase1_padronizar_extrato import executar as fase1_executar
from granularizar_sispag import executar as granularizar_sispag_executar
from leitura.contas_pagas import ler_contas_pagas, montar_pagamentos
from leitura.extrato_padronizado import ler_extrato_padronizado
from preencher_nova_descricao_sispag import executar as preencher_nova_descricao_executar
from validacao.validador import STATUS_APROVADO, STATUS_BLOQUEADO, determinar_competencia
from validar_saida_contabil import executar as validar_saida_contabil_executar

STATUS_TXT_GERADO = "TXT_GERADO_VALIDADO"
STATUS_BLOQUEADO_PARA_TXT = "BLOQUEADO_PARA_TXT"
STATUS_ERRO_PROCESSAMENTO = "ERRO_PROCESSAMENTO"

SUBPASTAS = ["intermediarios", "auditorias", "final", "logs"]

# Fases que dependem de cada fase anterior — usado só para marcar
# "NAO_EXECUTADA" de forma legível no relatório quando algo falha antes delas.
FASES_RESTANTES = {
    "2-processar-extrato": [
        "3-processar-contas-pagas", "4-conciliar-sispag", "5-6-nova-descricao-e-tarifas",
        "5b-nova-descricao-sispag", "7-de-para", "8-granularizar-sispag", "9-validacao-contabil", "10-exportar-txt",
    ],
    "3-processar-contas-pagas": [
        "4-conciliar-sispag", "5-6-nova-descricao-e-tarifas", "5b-nova-descricao-sispag",
        "7-de-para", "8-granularizar-sispag", "9-validacao-contabil", "10-exportar-txt",
    ],
    "4-conciliar-sispag": [
        "5b-nova-descricao-sispag", "7-de-para", "8-granularizar-sispag", "9-validacao-contabil", "10-exportar-txt",
    ],
    "5-6-nova-descricao-e-tarifas": [
        "5b-nova-descricao-sispag", "7-de-para", "8-granularizar-sispag", "9-validacao-contabil", "10-exportar-txt",
    ],
    "5b-nova-descricao-sispag": ["7-de-para", "8-granularizar-sispag", "9-validacao-contabil", "10-exportar-txt"],
    "7-de-para": ["8-granularizar-sispag", "9-validacao-contabil", "10-exportar-txt"],
    "8-granularizar-sispag": ["9-validacao-contabil", "10-exportar-txt"],
    "9-validacao-contabil": ["10-exportar-txt"],
}


@dataclass
class ResultadoFase:
    nome: str
    status: str  # OK | FALHA | BLOQUEADA | NAO_EXECUTADA
    detalhes: dict = field(default_factory=dict)
    erro: Optional[str] = None


@dataclass
class ResultadoPipeline:
    status: str
    competencia: Optional[str]
    resumo: dict
    arquivos_gerados: dict
    pendencias: list
    erros: list
    fases: list  # list[ResultadoFase]
    diretorio_saida: Optional[str]


class PipelineError(Exception):
    """Erro de validação de entrada do próprio orquestrador (não de uma fase)."""


def _preparar_diretorios(base: Path, competencia: str) -> dict[str, Path]:
    raiz = base / competencia
    caminhos = {nome: raiz / nome for nome in SUBPASTAS}
    for caminho in caminhos.values():
        caminho.mkdir(parents=True, exist_ok=True)
    caminhos["raiz"] = raiz
    return caminhos


def _proximo_diretorio_disponivel(base: Path, competencia: str) -> Path:
    """Nunca sobrescreve uma execução anterior da mesma competência: se
    'saida/2026-08' já existe, usa 'saida/2026-08_v2', '_v3'... em vez de
    apagar o que já estava lá."""
    candidato = base / competencia
    if not candidato.exists():
        return candidato
    versao = 2
    while (base / f"{competencia}_v{versao}").exists():
        versao += 1
    return base / f"{competencia}_v{versao}"


def _log(caminho_log: Path, fase: str, status: str, detalhes: dict, erro: Optional[str] = None) -> None:
    timestamp = datetime.now().isoformat(timespec="seconds")
    linha = f"{timestamp} | {fase} | {status} | {json.dumps(detalhes, ensure_ascii=False, default=str)}"
    if erro:
        linha += f" | ERRO: {erro}"
    with open(caminho_log, "a", encoding="utf-8") as f:
        f.write(linha + "\n")


def processar_mes(
    caminho_extrato: str,
    caminho_contas_pagas: str,
    caminho_de_para: str,
    caminho_correcoes_manuais: str,
    diretorio_saida_base: str = "saida",
) -> ResultadoPipeline:
    """Executa o processamento mensal completo, do extrato bruto ao TXT.

    Recebe explicitamente os 4 arquivos de entrada — nunca depende de
    "o que estiver na pasta". A competência de saída (ex.: '2026-08') é
    determinada a partir das próprias datas do extrato (Fase 2), nunca do
    nome dos arquivos de entrada.
    """
    fases: list[ResultadoFase] = []
    erros: list[str] = []
    pendencias: list[str] = []
    arquivos_gerados: dict[str, str] = {}
    resumo: dict = {}
    competencia: Optional[str] = None

    # --- validar arquivos de entrada ---------------------------------------
    for rotulo, caminho in [
        ("EXTRATO", caminho_extrato),
        ("CONTAS_PAGAS", caminho_contas_pagas),
        ("DE_PARA", caminho_de_para),
        ("CORRECOES_MANUAIS", caminho_correcoes_manuais),
    ]:
        if not Path(caminho).exists():
            erro = f"{rotulo} não encontrado: {caminho}"
            fases.append(ResultadoFase("1-validar-entradas", "FALHA", {}, erro))
            erros.append(erro)
            return ResultadoPipeline(
                STATUS_ERRO_PROCESSAMENTO, None, resumo, arquivos_gerados, pendencias, erros, fases, None
            )
    fases.append(ResultadoFase("1-validar-entradas", "OK", {"arquivos_confirmados": 4}))

    provisorio = f"_processando_{datetime.now().strftime('%Y%m%d%H%M%S')}"
    dirs = _preparar_diretorios(Path(diretorio_saida_base), provisorio)
    caminho_log = dirs["logs"] / "pipeline.log"
    _log(caminho_log, "1-validar-entradas", "OK", {"arquivos_confirmados": 4})

    def _falhar(nome_fase: str, exc: Exception) -> ResultadoPipeline:
        erro = f"{type(exc).__name__}: {exc}"
        fases.append(ResultadoFase(nome_fase, "FALHA", {}, erro))
        erros.append(erro)
        _log(caminho_log, nome_fase, "FALHA", {}, erro)
        for restante in FASES_RESTANTES.get(nome_fase, []):
            fases.append(ResultadoFase(restante, "NAO_EXECUTADA", {}))
            pendencias.append(f"{restante} não executada: fase anterior ({nome_fase}) falhou.")
        return ResultadoPipeline(
            STATUS_ERRO_PROCESSAMENTO, competencia, resumo, arquivos_gerados, pendencias, erros, fases,
            str(dirs["raiz"]),
        )

    # --- processar Extrato ---------------------------------------------------
    caminho_extrato_padronizado = dirs["intermediarios"] / "extrato_padronizado.xlsx"
    try:
        fase1_executar(caminho_extrato, str(caminho_extrato_padronizado))
        # fase1_executar devolve o DataFrame interno (colunas em minúsculo);
        # para determinar competência e contar movimentações, relê o Excel
        # já gravado com o leitor padrão do projeto (mesmas colunas usadas
        # pelo resto do pipeline).
        df_extrato = ler_extrato_padronizado(str(caminho_extrato_padronizado))
    except Exception as exc:
        return _falhar("2-processar-extrato", exc)

    competencia_periodo = determinar_competencia(df_extrato)
    competencia = competencia_periodo[0].strftime("%Y-%m") if competencia_periodo else "competencia-indeterminada"

    # A pasta provisória já tem tudo que a Fase 2 gravou — só renomeia para o
    # nome definitivo da competência, sem reprocessar nada. Se já existir uma
    # execução anterior da mesma competência, usa uma pasta versionada em vez
    # de apagar o que já estava lá.
    dirs_definitivo_raiz = _proximo_diretorio_disponivel(Path(diretorio_saida_base), competencia)
    if dirs_definitivo_raiz.name != competencia:
        pendencias.append(
            f"Já existia uma execução anterior para {competencia} — esta ficou em "
            f"'{dirs_definitivo_raiz.name}' para não sobrescrever a anterior."
        )
    shutil.move(str(dirs["raiz"]), str(dirs_definitivo_raiz))
    dirs = {nome: dirs_definitivo_raiz / nome for nome in SUBPASTAS}
    dirs["raiz"] = dirs_definitivo_raiz
    caminho_extrato_padronizado = dirs["intermediarios"] / "extrato_padronizado.xlsx"
    caminho_log = dirs["logs"] / "pipeline.log"

    arquivos_gerados["extrato_padronizado"] = str(caminho_extrato_padronizado)
    detalhes = {"movimentacoes": len(df_extrato)}
    fases.append(ResultadoFase("2-processar-extrato", "OK", detalhes))
    _log(caminho_log, "2-processar-extrato", "OK", detalhes)
    resumo["movimentacoes_bancarias"] = len(df_extrato)

    # --- processar Contas Pagas ----------------------------------------------
    try:
        df_contas_pagas = ler_contas_pagas(caminho_contas_pagas)
        candidatos, fora_do_escopo = montar_pagamentos(df_contas_pagas)
    except Exception as exc:
        return _falhar("3-processar-contas-pagas", exc)

    detalhes = {
        "linhas_origem": len(df_contas_pagas),
        "pagamentos_candidatos": len(candidatos),
        "fora_do_escopo": len(fora_do_escopo),
    }
    fases.append(ResultadoFase("3-processar-contas-pagas", "OK", detalhes))
    _log(caminho_log, "3-processar-contas-pagas", "OK", detalhes)
    resumo["pagamentos_candidatos"] = len(candidatos)

    # --- conciliar SISPAG ------------------------------------------------------
    caminho_auditoria_sispag = dirs["auditorias"] / "auditoria_sispag.xlsx"
    try:
        resultados_sispag = conciliar_sispag_executar(
            str(caminho_extrato_padronizado), caminho_contas_pagas, str(caminho_auditoria_sispag)
        )
    except Exception as exc:
        return _falhar("4-conciliar-sispag", exc)

    contagem_sispag = {STATUS_EXATO_1X1: 0, STATUS_EXATO_SOMA: 0, SISPAG_AMBIGUO: 0, SISPAG_NAO_ENCONTRADO: 0}
    for r in resultados_sispag:
        contagem_sispag[r.status] = contagem_sispag.get(r.status, 0) + 1

    arquivos_gerados["auditoria_sispag"] = str(caminho_auditoria_sispag)
    fases.append(ResultadoFase("4-conciliar-sispag", "OK", contagem_sispag))
    _log(caminho_log, "4-conciliar-sispag", "OK", contagem_sispag)
    resumo["sispag"] = contagem_sispag
    if contagem_sispag[SISPAG_AMBIGUO] or contagem_sispag[SISPAG_NAO_ENCONTRADO]:
        pendencias.append(
            f"{contagem_sispag[SISPAG_AMBIGUO]} SISPAG AMBIGUO e "
            f"{contagem_sispag[SISPAG_NAO_ENCONTRADO]} NAO_ENCONTRADO precisam de revisão manual "
            "(ver auditoria_sispag.xlsx)."
        )

    # --- NOVA DESCRICAO (comuns + regras fixas de tarifa) ---------------------
    caminho_saida_enriquecida = dirs["intermediarios"] / "saida_enriquecida.xlsx"
    try:
        linhas_enriquecidas = enriquecer_extrato_executar(
            str(caminho_extrato_padronizado), str(caminho_saida_enriquecida)
        )
    except Exception as exc:
        return _falhar("5-6-nova-descricao-e-tarifas", exc)

    qtd_tarifas = sum(1 for l in linhas_enriquecidas if l["_eh_tarifa"])
    detalhes = {"linhas": len(linhas_enriquecidas), "tarifas_classificadas": qtd_tarifas}
    arquivos_gerados["saida_enriquecida"] = str(caminho_saida_enriquecida)
    fases.append(ResultadoFase("5-6-nova-descricao-e-tarifas", "OK", detalhes))
    _log(caminho_log, "5-6-nova-descricao-e-tarifas", "OK", detalhes)
    resumo["tarifas_classificadas"] = qtd_tarifas

    # --- NOVA DESCRICAO dos SISPAG (usa auditoria_sispag) ----------------------
    caminho_saida_enriquecida_sispag = dirs["intermediarios"] / "saida_enriquecida_sispag.xlsx"
    try:
        resultados_merge = preencher_nova_descricao_executar(
            str(caminho_saida_enriquecida), str(caminho_auditoria_sispag), str(caminho_saida_enriquecida_sispag)
        )
    except Exception as exc:
        return _falhar("5b-nova-descricao-sispag", exc)

    preenchidos = sum(1 for r in resultados_merge if r.preenchido)
    detalhes = {"sispag_processados": len(resultados_merge), "preenchidos": preenchidos}
    arquivos_gerados["saida_enriquecida_sispag"] = str(caminho_saida_enriquecida_sispag)
    fases.append(ResultadoFase("5b-nova-descricao-sispag", "OK", detalhes))
    _log(caminho_log, "5b-nova-descricao-sispag", "OK", detalhes)
    if preenchidos < len(resultados_merge):
        pendencias.append(
            f"{len(resultados_merge) - preenchidos} SISPAG não tiveram NOVA_DESCRICAO preenchida automaticamente."
        )

    # --- executar DE-PARA --------------------------------------------------------
    caminho_auditoria_de_para = dirs["auditorias"] / "auditoria_de_para.xlsx"
    try:
        linhas_de_para = executar_de_para_executar(
            str(caminho_saida_enriquecida_sispag), caminho_de_para, str(caminho_auditoria_de_para)
        )
    except Exception as exc:
        return _falhar("7-de-para", exc)

    contagem_de_para = dict(Counter(l["STATUS_DE_PARA"] for l in linhas_de_para))
    arquivos_gerados["auditoria_de_para"] = str(caminho_auditoria_de_para)
    fases.append(ResultadoFase("7-de-para", "OK", contagem_de_para))
    _log(caminho_log, "7-de-para", "OK", contagem_de_para)
    resumo["de_para"] = contagem_de_para
    if contagem_de_para.get("SEM_DE_PARA") or contagem_de_para.get("AMBIGUO"):
        pendencias.append(
            f"{contagem_de_para.get('SEM_DE_PARA', 0)} SEM_DE_PARA e "
            f"{contagem_de_para.get('AMBIGUO', 0)} AMBIGUO precisam de correção manual antes da validação final."
        )

    # --- granularizar SISPAG -------------------------------------------------------
    caminho_saida_contabil = dirs["intermediarios"] / "saida_contabil.xlsx"
    caminho_auditoria_granularizacao = dirs["auditorias"] / "auditoria_granularizacao.xlsx"
    try:
        linhas_contabeis, auditorias_gran, requer_class = granularizar_sispag_executar(
            str(caminho_extrato_padronizado),
            caminho_contas_pagas,
            str(caminho_auditoria_sispag),
            str(caminho_auditoria_de_para),
            caminho_de_para,
            str(caminho_saida_contabil),
            str(caminho_auditoria_granularizacao),
        )
    except Exception as exc:
        return _falhar("8-granularizar-sispag", exc)

    detalhes = {"linhas_contabeis": len(linhas_contabeis), "requer_granularizacao_classificacao": len(requer_class)}
    arquivos_gerados["saida_contabil"] = str(caminho_saida_contabil)
    arquivos_gerados["auditoria_granularizacao"] = str(caminho_auditoria_granularizacao)
    fases.append(ResultadoFase("8-granularizar-sispag", "OK", detalhes))
    _log(caminho_log, "8-granularizar-sispag", "OK", detalhes)
    resumo["linhas_contabeis"] = len(linhas_contabeis)

    # --- executar validação contábil ------------------------------------------------
    caminho_relatorio_validacao = dirs["auditorias"] / "relatorio_validacao.xlsx"
    try:
        saida_validada, status_validacao = validar_saida_contabil_executar(
            str(caminho_saida_contabil),
            caminho_correcoes_manuais,
            str(caminho_auditoria_granularizacao),
            str(caminho_relatorio_validacao),
        )
    except Exception as exc:
        return _falhar("9-validacao-contabil", exc)

    qtd_aprovadas = int((saida_validada["STATUS_VALIDACAO"] == STATUS_APROVADO).sum())
    qtd_bloqueadas = int((saida_validada["STATUS_VALIDACAO"] == STATUS_BLOQUEADO).sum())
    detalhes = {"status_geral": status_validacao, "aprovadas": qtd_aprovadas, "bloqueadas": qtd_bloqueadas}
    arquivos_gerados["relatorio_validacao"] = str(caminho_relatorio_validacao)
    fases.append(ResultadoFase("9-validacao-contabil", "OK", detalhes))
    _log(caminho_log, "9-validacao-contabil", "OK", detalhes)
    resumo["linhas_aprovadas"] = qtd_aprovadas
    resumo["linhas_bloqueadas"] = qtd_bloqueadas

    # --- gerar TXT — só se aprovado -----------------------------------------------
    if status_validacao != STATUS_GERAL_APROVADO:
        fases.append(
            ResultadoFase(
                "10-exportar-txt", "BLOQUEADA",
                {"motivo": f"STATUS_GERAL={status_validacao}, não é {STATUS_GERAL_APROVADO}"},
            )
        )
        _log(caminho_log, "10-exportar-txt", "BLOQUEADA", {"status_validacao": status_validacao})
        pendencias.append(
            f"TXT não gerado: {qtd_bloqueadas} linha(s) BLOQUEADA(s) em relatorio_validacao.xlsx — corrija e reprocesse."
        )
        status_final = STATUS_BLOQUEADO_PARA_TXT
    else:
        caminho_txt = dirs["final"] / f"{competencia}_importacao_contabil.txt"
        caminho_auditoria_exportacao = dirs["auditorias"] / "auditoria_exportacao_txt.xlsx"
        try:
            status_exportacao, linhas_exportadas, problemas_exportacao = exportacao_txt_executar(
                str(caminho_relatorio_validacao), str(caminho_txt), str(caminho_auditoria_exportacao)
            )
        except Exception as exc:
            return _falhar("10-exportar-txt", exc)

        total = sum((l.valor_original for l in linhas_exportadas), start=Decimal("0.00"))
        detalhes = {
            "status": status_exportacao, "lancamentos": len(linhas_exportadas),
            "total": str(total), "problemas_round_trip": len(problemas_exportacao),
        }
        arquivos_gerados["txt"] = str(caminho_txt)
        arquivos_gerados["auditoria_exportacao_txt"] = str(caminho_auditoria_exportacao)
        fases.append(
            ResultadoFase("10-exportar-txt", "OK" if status_exportacao == STATUS_TXT_VALIDADO else "FALHA", detalhes)
        )
        _log(caminho_log, "10-exportar-txt", status_exportacao, detalhes)
        resumo["total_exportado"] = str(total)
        resumo["lancamentos_exportados"] = len(linhas_exportadas)

        if status_exportacao != STATUS_TXT_VALIDADO:
            erros.append(f"Exportação gerou TXT mas pós-validação encontrou {len(problemas_exportacao)} problema(s).")
            status_final = STATUS_ERRO_PROCESSAMENTO
        else:
            status_final = STATUS_TXT_GERADO

    # --- resumo final -----------------------------------------------------------------
    fases.append(ResultadoFase("11-resumo-final", "OK", {"status_final": status_final}))
    _log(caminho_log, "11-resumo-final", "OK", {"status_final": status_final})

    resultado = ResultadoPipeline(
        status=status_final,
        competencia=competencia,
        resumo=resumo,
        arquivos_gerados=arquivos_gerados,
        pendencias=pendencias,
        erros=erros,
        fases=fases,
        diretorio_saida=str(dirs["raiz"]),
    )

    caminho_resumo = dirs["logs"] / "resumo_execucao.json"
    with open(caminho_resumo, "w", encoding="utf-8") as f:
        json.dump(
            {
                "status": resultado.status, "competencia": resultado.competencia, "resumo": resultado.resumo,
                "arquivos_gerados": resultado.arquivos_gerados, "pendencias": resultado.pendencias,
                "erros": resultado.erros,
            },
            f, ensure_ascii=False, indent=2, default=str,
        )
    arquivos_gerados["resumo_execucao"] = str(caminho_resumo)

    return resultado
