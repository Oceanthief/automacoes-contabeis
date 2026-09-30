"""
Motor de classificação contábil (Fase 3).

Consome os `Movimento` já lidos e normalizados pela Fase 2
(`src/leitura_extrato.py`) e o DE-PARA oficial (planilha de controle de
fornecedores e clientes) para produzir uma `Classificacao` por
movimento — nunca mais, nunca menos.

Fora de escopo desta fase (ver enunciado): geração de Excel final ou de
TXT de importação. Este módulo só classifica; não grava nada.

--------------------------------------------------------------------------
DECISÕES DE NEGÓCIO REGISTRADAS
--------------------------------------------------------------------------
O primeiro rascunho da especificação da Fase 3 contradizia 3 decisões já
confirmadas pela contabilidade no contrato de dados (Fase 1). Prevaleceram
as decisões confirmadas:

  1. Transferência SC/Débito: `débito=11221 / crédito=11037` (não é
     mais `REGRA_SC_DEBITO_PENDENTE`).
  2. Salário: tolerância de ±3 dias em torno dos dias 6 e 20 (não é
     mais só o dia exato).
  3. Pró-labore: sempre prevalece sobre a regra de salário, mesmo
     quando o histórico contém os dois gatilhos ao mesmo tempo.

Continua em aberto: o lado CRÉDITO da transferência de SAÍDA para a
filial PR — só o lado débito foi confirmado (é o único observado nos
dados). Se um dia aparecer em crédito, cai em `REGRA_PENDENTE`.
"""

from __future__ import annotations

import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "config"))
from leitura_extrato import Movimento, _normalizar_texto  # reaproveita a mesma normalização da Fase 2
import configuracoes as cfg


# ---------------------------------------------------------------------------
# Constantes de status (exatamente os 6 estados aceitos pelo enunciado)
# ---------------------------------------------------------------------------

STATUS_OK = "OK"
STATUS_PENDENTE_MANUAL = "PENDENTE_MANUAL"
STATUS_NAO_ENCONTRADO = "NAO_ENCONTRADO"
STATUS_AMBIGUO = "AMBIGUO"
STATUS_REGRA_PENDENTE = "REGRA_PENDENTE"
STATUS_DADO_INVALIDO = "DADO_INVALIDO"

_PREENCHER_MANUAL = cfg.MARCADOR_PREENCHER_MANUAL

_TRANSF_PR = cfg.TRANSF_ENTRADA_PR_TEXTO
_TRANSF_SC = cfg.TRANSF_SC_TEXTO
# Confirmado pela contabilidade: quando esta frase (transferência de SAÍDA para a
# conta PR) aparece com valor em DÉBITO, débito=11023 / crédito=11037 —
# mesma conta que a Obs 3 do Controle já define para "DA CONTA...PR" no
# lado débito. O lado crédito desta frase nunca foi observado nos dados
# nem confirmado, então continua pendente se aparecer.
_TRANSF_PR_SAIDA_NAO_DEFINIDA = cfg.TRANSF_SAIDA_PR_TEXTO
_SALARIO_GATILHO = cfg.SALARIO_GATILHO_TEXTO

# Aliases que NÃO existem na planilha Controle, mas foram confirmados pela
# contabilidade (ex.: o banco escreve "J P" onde o cadastro tem "J.P.").
# Enquanto o Controle não for atualizado, o alias é aplicado aqui, de forma
# explícita e auditável — nunca escondida dentro do matching normal — para
# uma correspondência já confirmada não virar NAO_ENCONTRADO.
ALIASES_CONFIRMADOS: dict[str, str] = cfg.ALIASES_DEPARA


# ---------------------------------------------------------------------------
# DE-PARA
# ---------------------------------------------------------------------------

@dataclass
class DeParaEntrada:
    nome_original: str
    nome_normalizado: str
    conta_debito: str  # sempre string ("PREENCHER MANUAL" incluso), nunca inventado
    conta_credito: str
    linha_origem: int  # linha na planilha Controle, para auditoria
    via_alias: bool = False  # True quando veio de ALIASES_CONFIRMADOS, não do arquivo


def carregar_depara(caminho: str | Path, aplicar_aliases: bool = True) -> list[DeParaEntrada]:
    """Lê a planilha Controle (fornecedores e clientes) como está.

    Não interpreta nem corrige nada do conteúdo — só carrega. Contas são
    sempre mantidas como string (inclusive `PREENCHER MANUAL` literal).
    """
    wb = openpyxl.load_workbook(str(caminho), data_only=True)
    ws = wb["Planilha1"]

    entradas: list[DeParaEntrada] = []
    indice_por_nome: dict[str, DeParaEntrada] = {}
    for i, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True), start=2):
        nome = row[0]
        if nome is None:
            continue
        conta_debito = str(row[1]).strip() if row[1] is not None else ""
        conta_credito = str(row[2]).strip() if row[2] is not None else ""
        entrada = DeParaEntrada(
            nome_original=str(nome).strip(),
            nome_normalizado=_normalizar_texto(nome),
            conta_debito=conta_debito,
            conta_credito=conta_credito,
            linha_origem=i,
        )
        entradas.append(entrada)
        indice_por_nome.setdefault(entrada.nome_normalizado, entrada)

    if aplicar_aliases:
        for alias, canonico in ALIASES_CONFIRMADOS.items():
            alias_norm = _normalizar_texto(alias)
            canonico_norm = _normalizar_texto(canonico)
            base = indice_por_nome.get(canonico_norm)
            if base is None:
                # o cadastro canônico não existe no arquivo — não inventamos
                # o alias sem uma base real para apontar.
                continue
            entradas.append(
                DeParaEntrada(
                    nome_original=f"{alias} (alias confirmado de: {base.nome_original})",
                    nome_normalizado=alias_norm,
                    conta_debito=base.conta_debito,
                    conta_credito=base.conta_credito,
                    linha_origem=base.linha_origem,
                    via_alias=True,
                )
            )

    return entradas


def _eh_preencher_manual(conta: str) -> bool:
    return _normalizar_texto(conta) == _PREENCHER_MANUAL


@dataclass
class ResultadoBusca:
    vencedoras: list[DeParaEntrada]  # 0, 1, ou várias (se ambíguo de verdade)
    empatadas_mas_iguais: bool  # True quando há >1 vencedora mas todas concordam no resultado


def _buscar_depara(hist_normalizado: str, depara: list[DeParaEntrada]) -> ResultadoBusca:
    """Encontra, entre as entradas do De-Para, as que "cabem" no histórico.

    Matching determinístico por substring (nome do De-Para contido no
    histórico normalizado) — sem fuzzy matching, como pedido. Em caso de
    mais de uma correspondência, vence a mais longa (mais específica) —
    essa estratégia já foi validada contra julho e agosto na Fase 1
    (ex.: "BANCO SANTANDER" vs. "BANCO SANTANDER - CONSORCIO").

    Duas entradas empatadas em comprimento que levam à MESMA conta
    (débito e crédito iguais) não são tratadas como ambíguas — são
    duplicatas inofensivas do cadastro (ex.: "BANCO SANTANDER" cadastrado
    duas vezes com os mesmos valores). Só vira `AMBIGUO` quando o empate
    leva a contas diferentes.
    """
    candidatas = [e for e in depara if e.nome_normalizado and e.nome_normalizado in hist_normalizado]
    if not candidatas:
        return ResultadoBusca(vencedoras=[], empatadas_mas_iguais=False)

    maior_tamanho = max(len(e.nome_normalizado) for e in candidatas)
    vencedoras = [e for e in candidatas if len(e.nome_normalizado) == maior_tamanho]

    contas_distintas = {(e.conta_debito, e.conta_credito) for e in vencedoras}
    empatadas_mas_iguais = len(vencedoras) > 1 and len(contas_distintas) == 1

    return ResultadoBusca(vencedoras=vencedoras, empatadas_mas_iguais=empatadas_mas_iguais)


# ---------------------------------------------------------------------------
# Resultado da classificação
# ---------------------------------------------------------------------------

@dataclass
class Classificacao:
    source_row: int
    historico_original: str
    chave_encontrada: Optional[str]
    regra_aplicada: str
    conta_debito: Optional[str]
    conta_credito: Optional[str]
    status: str
    detalhes: str = ""
    flags_fase2: list[str] = field(default_factory=list)

    def to_row(self) -> dict:
        return {
            "SOURCE_ROW": self.source_row,
            "HISTORICO_ORIGINAL": self.historico_original,
            "CHAVE_ENCONTRADA": self.chave_encontrada,
            "REGRA_APLICADA": self.regra_aplicada,
            "CONTA_DEBITO": self.conta_debito,
            "CONTA_CREDITO": self.conta_credito,
            "STATUS": self.status,
            "DETALHES": self.detalhes,
        }


# ---------------------------------------------------------------------------
# Regras específicas
# ---------------------------------------------------------------------------

def _regra_transferencias(mov: Movimento) -> Optional[Classificacao]:
    hist = mov.historico_normalizado

    if hist == _normalizar_texto(_TRANSF_PR):
        if mov.credito is not None:
            return Classificacao(mov.source_row, mov.historico_original, _TRANSF_PR,
                                  "transferencia_pr_credito", cfg.CONTA_SANTANDER, cfg.CONTA_TRANSFERENCIA_PR, STATUS_OK)
        else:  # débito preenchido — não observado em julho/agosto, mas coberto pela Obs 3 do Controle
            return Classificacao(mov.source_row, mov.historico_original, _TRANSF_PR,
                                  "transferencia_pr_debito", cfg.CONTA_TRANSFERENCIA_PR, cfg.CONTA_SANTANDER, STATUS_OK)

    if hist == _normalizar_texto(_TRANSF_PR_SAIDA_NAO_DEFINIDA):
        # Confirmado pela contabilidade: mesma lógica da frase "DA CONTA...PR"
        # no lado débito. Nos dados reais essa frase só aparece em débito,
        # mas isso não é presumido para o lado crédito, que nunca foi
        # confirmado nem observado — cai em REGRA_PENDENTE se aparecer.
        if mov.debito is not None:
            return Classificacao(mov.source_row, mov.historico_original, _TRANSF_PR_SAIDA_NAO_DEFINIDA,
                                  "transferencia_pr_saida_debito", cfg.CONTA_TRANSFERENCIA_PR, cfg.CONTA_SANTANDER, STATUS_OK)
        else:
            return Classificacao(mov.source_row, mov.historico_original, _TRANSF_PR_SAIDA_NAO_DEFINIDA,
                                  "transferencia_pr_saida_credito_pendente", None, None, STATUS_REGRA_PENDENTE,
                                  detalhes="A conta desta frase só foi confirmada no lado DÉBITO "
                                           f"(débito={cfg.CONTA_TRANSFERENCIA_PR}/crédito={cfg.CONTA_SANTANDER}). "
                                           "O lado CRÉDITO nunca apareceu nos dados e não foi confirmado — "
                                           "o mesmo valor não é presumido.")

    if hist == _normalizar_texto(_TRANSF_SC):
        if mov.credito is not None:
            return Classificacao(mov.source_row, mov.historico_original, _TRANSF_SC,
                                  "transferencia_sc_credito", cfg.CONTA_SANTANDER, cfg.CONTA_TRANSFERENCIA_SC, STATUS_OK)
        else:
            # Confirmado pela contabilidade na Fase 1 (Ambiguidade A2) e
            # reconfirmado quando a especificação da Fase 3 divergiu disso.
            return Classificacao(mov.source_row, mov.historico_original, _TRANSF_SC,
                                  "transferencia_sc_debito", cfg.CONTA_TRANSFERENCIA_SC, cfg.CONTA_SANTANDER, STATUS_OK)
    return None


def _regra_salarios(mov: Movimento) -> Optional[Classificacao]:
    if _normalizar_texto(_SALARIO_GATILHO) not in mov.historico_normalizado:
        return None

    dia = mov.data.day  # seguro: DADO_INVALIDO já filtrou movimentos sem data antes de chegar aqui

    # Tolerância de ±N dias em torno de cada âncora, confirmada pela
    # contabilidade na Fase 1 (Ambiguidade A1). As janelas não se sobrepõem (âncoras a
    # 14 dias de distância, tolerância de 3), então não há ambiguidade
    # entre as duas.
    if abs(dia - cfg.SALARIO_ANCORA_1) <= cfg.SALARIO_TOLERANCIA_DIAS:
        # crédito=11037 vem da Regra 1 (PAGO) já confirmada no contrato —
        # o histórico real sempre começa com "Pago:".
        return Classificacao(mov.source_row, mov.historico_original, _SALARIO_GATILHO,
                              "salario_dia_06_tolerancia", cfg.CONTA_SALARIO_DIA_06, cfg.CONTA_SANTANDER, STATUS_OK,
                              detalhes=f"Dia do mês = {dia} (âncora: dia {cfg.SALARIO_ANCORA_1}, "
                                       f"tolerância ±{cfg.SALARIO_TOLERANCIA_DIAS} dias).")
    if abs(dia - cfg.SALARIO_ANCORA_2) <= cfg.SALARIO_TOLERANCIA_DIAS:
        return Classificacao(mov.source_row, mov.historico_original, _SALARIO_GATILHO,
                              "salario_dia_20_tolerancia", cfg.CONTA_SALARIO_DIA_20, cfg.CONTA_SANTANDER, STATUS_OK,
                              detalhes=f"Dia do mês = {dia} (âncora: dia {cfg.SALARIO_ANCORA_2}, "
                                       f"tolerância ±{cfg.SALARIO_TOLERANCIA_DIAS} dias).")

    # Fora de qualquer uma das duas janelas de tolerância — isso não tem
    # decisão confirmada (nem a tolerância de ±3 dias cobre esse caso).
    return Classificacao(mov.source_row, mov.historico_original, _SALARIO_GATILHO,
                          "salario_fora_da_tolerancia", None, None, STATUS_REGRA_PENDENTE,
                          detalhes=f"Dia do mês = {dia}, fora da janela de ±3 dias em torno dos dias 6 e "
                                   "20 (a única tolerância confirmada até agora).")


def _regra_pro_labore_prioritario(mov: Movimento, depara: list[DeParaEntrada]) -> Optional[Classificacao]:
    """Confirmado pela contabilidade na Fase 1 (Ambiguidade A10): sempre que
    o histórico contiver "PRO LABORE", essa classificação prevalece sobre a
    regra de salário — mesmo quando o mesmo histórico também contém o
    gatilho de salário (ex.: "...SALARIOS- EMPRESA - PRO LABORE SOCIO A").
    Por isso este passo roda ANTES da regra de salários na ordem de
    precedência.

    Reaproveita a mesma busca do DE-PARA normal (passo 5) — o Controle
    pode ter entradas específicas por sócio ("PRO LABORE SOCIO A") e uma
    genérica ("PRO LABORE") que cobre qualquer variação futura.
    """
    if cfg.PRO_LABORE_GATILHO_TEXTO not in mov.historico_normalizado:
        return None
    resultado = _regra_depara_normal(mov, depara)
    resultado.regra_aplicada = f"pro_labore_prioritario:{resultado.regra_aplicada}"
    return resultado


def _regra_pago_ou_recebido(mov: Movimento, depara: list[DeParaEntrada]) -> Optional[Classificacao]:
    hist = mov.historico_normalizado

    if hist.startswith("PAGO"):
        prefixo, conta_fixa_lado = "pago", "credito"
    elif hist.startswith("RECEBIDO"):
        prefixo, conta_fixa_lado = "recebido", "debito"
    else:
        return None

    busca = _buscar_depara(hist, depara)

    if not busca.vencedoras:
        return Classificacao(mov.source_row, mov.historico_original, None,
                              f"{prefixo}_sem_depara", None, None, STATUS_NAO_ENCONTRADO,
                              detalhes="Histórico começa com "
                                       f"{'PAGO' if prefixo == 'pago' else 'RECEBIDO'}, mas nenhum nome do "
                                       "De-Para foi encontrado dentro do texto.")

    if len(busca.vencedoras) > 1 and not busca.empatadas_mas_iguais:
        nomes = ", ".join(sorted(e.nome_original for e in busca.vencedoras))
        return Classificacao(mov.source_row, mov.historico_original, nomes,
                              f"{prefixo}_depara_ambiguo", None, None, STATUS_AMBIGUO,
                              detalhes=f"Mais de uma chave do De-Para (com contas diferentes) cabe neste "
                                       f"histórico, no mesmo comprimento máximo: {nomes}.")

    entrada = busca.vencedoras[0]
    conta_variavel = entrada.conta_debito if conta_fixa_lado == "credito" else entrada.conta_credito

    if _eh_preencher_manual(conta_variavel):
        return Classificacao(mov.source_row, mov.historico_original, entrada.nome_original,
                              f"{prefixo}_depara_manual", None, None, STATUS_PENDENTE_MANUAL,
                              detalhes=f"Conta de {'débito' if conta_fixa_lado == 'credito' else 'crédito'} "
                                       f"cadastrada como PREENCHER MANUAL para '{entrada.nome_original}'.")

    if conta_fixa_lado == "credito":
        return Classificacao(mov.source_row, mov.historico_original, entrada.nome_original,
                              f"{prefixo}_depara", conta_variavel, cfg.CONTA_SANTANDER, STATUS_OK)
    else:
        return Classificacao(mov.source_row, mov.historico_original, entrada.nome_original,
                              f"{prefixo}_depara", cfg.CONTA_SANTANDER, conta_variavel, STATUS_OK)


def _regra_depara_normal(mov: Movimento, depara: list[DeParaEntrada]) -> Classificacao:
    """Passo 5 — DE-PARA direto, para históricos que não começam com
    PAGO/RECEBIDO e não são transferência nem salário (ex.: tarifas,
    IOF, FGTS quando vier sem o prefixo "Pago:", pró-labore avulso)."""
    hist = mov.historico_normalizado
    busca = _buscar_depara(hist, depara)

    if not busca.vencedoras:
        return Classificacao(mov.source_row, mov.historico_original, None,
                              "depara_nao_encontrado", None, None, STATUS_NAO_ENCONTRADO,
                              detalhes="Nenhuma chave do De-Para foi encontrada dentro do histórico.")

    if len(busca.vencedoras) > 1 and not busca.empatadas_mas_iguais:
        nomes = ", ".join(sorted(e.nome_original for e in busca.vencedoras))
        return Classificacao(mov.source_row, mov.historico_original, nomes,
                              "depara_ambiguo", None, None, STATUS_AMBIGUO,
                              detalhes=f"Mais de uma chave do De-Para (com contas diferentes) cabe neste "
                                       f"histórico, no mesmo comprimento máximo: {nomes}.")

    entrada = busca.vencedoras[0]

    if _eh_preencher_manual(entrada.conta_debito) or _eh_preencher_manual(entrada.conta_credito):
        lado = "débito" if _eh_preencher_manual(entrada.conta_debito) else "crédito"
        return Classificacao(mov.source_row, mov.historico_original, entrada.nome_original,
                              "depara_manual", None, None, STATUS_PENDENTE_MANUAL,
                              detalhes=f"Conta de {lado} cadastrada como PREENCHER MANUAL para "
                                       f"'{entrada.nome_original}'.")

    return Classificacao(mov.source_row, mov.historico_original, entrada.nome_original,
                          "depara_direto", entrada.conta_debito, entrada.conta_credito, STATUS_OK)


# ---------------------------------------------------------------------------
# Função principal
# ---------------------------------------------------------------------------

def classificar_movimento(mov: Movimento, depara: list[DeParaEntrada]) -> Classificacao:
    """Classifica um único movimento, seguindo a ordem de precedência:

        1. transferências especiais (PR, SC)
        2. pró-labore (sempre vence sobre salário, mesmo se o histórico
           também contiver o gatilho de salário — confirmado na Fase 1)
        3. salários
        4. prefixo PAGO
        5. prefixo RECEBIDO
        6. DE-PARA normal
        7. pendência (NAO_ENCONTRADO, vindo do passo 6)

    Antes de tudo: se a Fase 2 já sinalizou o movimento como com dado
    estrutural inválido para virar um lançamento (data ausente, sem
    valor, débito e crédito simultâneos, ou histórico ausente), a
    classificação para aqui com `DADO_INVALIDO` — nenhuma regra de
    negócio tenta "adivinhar" um lançamento sem os dados mínimos que o
    próprio layout de saída exige (DATA, HISTORICO, VALOR).
    `valor_zero_*` e `linha_duplicada` NÃO bloqueiam a classificação —
    são informativos, não impedem determinar a conta.
    """
    flags_bloqueantes = {"data_invalida", "sem_valor", "debito_e_credito_preenchidos", "historico_ausente"}
    flags_presentes_bloqueantes = [f for f in mov.flags if f in flags_bloqueantes]
    if flags_presentes_bloqueantes:
        return Classificacao(
            mov.source_row, mov.historico_original, None, "dado_invalido_fase2",
            None, None, STATUS_DADO_INVALIDO,
            detalhes=f"Bloqueado por flag(s) da Fase 2: {flags_presentes_bloqueantes}.",
            flags_fase2=list(mov.flags),
        )

    for regra in (
        lambda: _regra_transferencias(mov),
        lambda: _regra_pro_labore_prioritario(mov, depara),
        lambda: _regra_salarios(mov),
        lambda: _regra_pago_ou_recebido(mov, depara),
    ):
        resultado = regra()
        if resultado is not None:
            resultado.flags_fase2 = list(mov.flags)
            return resultado

    resultado = _regra_depara_normal(mov, depara)
    resultado.flags_fase2 = list(mov.flags)
    return resultado


def classificar_movimentos(movimentos: list[Movimento], depara: list[DeParaEntrada]) -> list[Classificacao]:
    """Classifica uma lista inteira. Garante 1:1 — nenhum movimento some."""
    resultado = [classificar_movimento(m, depara) for m in movimentos]
    assert len(resultado) == len(movimentos), "nenhum movimento pode desaparecer na classificação"
    return resultado
