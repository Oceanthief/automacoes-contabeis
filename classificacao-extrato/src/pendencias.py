"""
Fluxo de pendências e conferência manual (Fase 4).

Gera `output/pendencias_<empresa>_AAAA_MM.xlsx` a partir do resultado da
Fase 3 (classificação), permite que um humano preencha CONTA_DEBITO_MANUAL
/ CONTA_CREDITO_MANUAL / OBSERVACAO_MANUAL (e, quando a data do movimento
for o problema, DATA_MANUAL), e reaplica esses ajustes num reprocessamento
posterior — sem nunca alterar a planilha Controle oficial nem o .xls de
origem.

Fora de escopo desta fase: gerar o TXT de importação (isso é Fase 5).

--------------------------------------------------------------------------
Como um item de pendência é considerado resolvido
--------------------------------------------------------------------------
Um item só sai da lista de bloqueios quando um humano preenche, no
arquivo de pendências, as DUAS contas (`CONTA_DEBITO_MANUAL` e
`CONTA_CREDITO_MANUAL`). Nunca inferimos uma a partir da outra, e nunca
reaplicamos regra automática sobre o valor ajustado — o preenchimento
manual É a decisão final, não um novo dado para o motor de regras
processar de novo.

Caso especial: quando o motivo do bloqueio inclui `data_invalida` (Fase
2), a linha também exige `DATA_MANUAL` preenchida, porque a data é campo
obrigatório do lançamento final e não pode ficar ausente mesmo com as
contas definidas.

Casos bloqueados por `sem_valor`, `debito_e_credito_preenchidos` ou
`historico_ausente` (Fase 2) só deveriam, no mundo ideal, ser corrigidos
no extrato de origem (que é somente leitura) e reprocessados desde a
Fase 2 — mas nada nesta camada impede o preenchimento manual das contas
como uma decisão consciente do conferente; o `OBSERVACAO_MANUAL` fica
como registro do porquê.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from datetime import date, datetime
from difflib import get_close_matches
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "config"))
import configuracoes as cfg
from leitura_extrato import Movimento
from classificador import (
    Classificacao,
    DeParaEntrada,
    STATUS_OK,
    STATUS_PENDENTE_MANUAL,
    STATUS_NAO_ENCONTRADO,
    STATUS_AMBIGUO,
    STATUS_REGRA_PENDENTE,
    STATUS_DADO_INVALIDO,
)

STATUS_BLOQUEANTES = {
    STATUS_NAO_ENCONTRADO,
    STATUS_PENDENTE_MANUAL,
    STATUS_AMBIGUO,
    STATUS_REGRA_PENDENTE,
    STATUS_DADO_INVALIDO,
}

MENSAGEM_BLOQUEIO = "PROCESSAMENTO NÃO LIBERADO PARA FINALIZAÇÃO"

_COLUNAS = [
    "SOURCE_ROW", "DATA", "HISTORICO_ORIGINAL", "DEBITO", "CREDITO",
    "NOME_IDENTIFICADO", "STATUS", "MOTIVO",
    "CONTA_DEBITO_MANUAL", "CONTA_CREDITO_MANUAL",
    "DATA_MANUAL", "OBSERVACAO_MANUAL", "SUGESTAO_DEPARA_AJUDA",
]
_COLUNAS_MANUAIS = {"CONTA_DEBITO_MANUAL", "CONTA_CREDITO_MANUAL", "DATA_MANUAL", "OBSERVACAO_MANUAL"}

_FONTE = "Arial"


def nome_arquivo_pendencias(ano: int, mes: int) -> str:
    return f"pendencias_{cfg.PREFIXO_ARQUIVOS}_{ano:04d}_{mes:02d}.xlsx"


# ---------------------------------------------------------------------------
# Ajuste manual
# ---------------------------------------------------------------------------

@dataclass
class AjusteManual:
    conta_debito: Optional[str] = None
    conta_credito: Optional[str] = None
    data: Optional[date] = None
    observacao: Optional[str] = None

    @property
    def contas_completas(self) -> bool:
        return bool(self.conta_debito) and bool(self.conta_credito)


def _texto_ou_none(valor) -> Optional[str]:
    if valor is None:
        return None
    texto = str(valor).strip()
    return texto or None


def carregar_ajustes_manuais(caminho: str | Path) -> dict[int, AjusteManual]:
    """Lê um arquivo de pendências já existente e extrai o que o humano
    preencheu — completo ou parcial, tanto o que ainda está aberto (aba
    "Pendencias") quanto o que já foi resolvido em execuções anteriores
    (aba "Resolvidos_Manualmente", que funciona como memória permanente
    do ajuste — uma vez resolvida, uma linha não pode "esquecer" a
    decisão do conferente só porque saiu da lista de pendências).
    Nunca falha se o arquivo não existir (primeira execução do mês):
    retorna dicionário vazio.
    """
    caminho = Path(caminho)
    if not caminho.exists():
        return {}

    wb = openpyxl.load_workbook(caminho, data_only=True)
    ajustes: dict[int, AjusteManual] = {}

    if "Resolvidos_Manualmente" in wb.sheetnames:
        ajustes.update(_ler_ajustes_de_aba(wb["Resolvidos_Manualmente"]))
    if "Pendencias" in wb.sheetnames:
        ajustes.update(_ler_ajustes_de_aba(wb["Pendencias"]))  # itens ainda abertos podem sobrescrever

    return ajustes


def _ler_ajustes_de_aba(ws) -> dict[int, AjusteManual]:
    cabecalho = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    if "SOURCE_ROW" not in cabecalho or "CONTA_DEBITO_MANUAL" not in cabecalho:
        return {}
    idx = {nome: i for i, nome in enumerate(cabecalho)}

    ajustes: dict[int, AjusteManual] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[idx["SOURCE_ROW"]] is None:
            continue
        source_row = int(row[idx["SOURCE_ROW"]])

        data_manual_bruta = row[idx["DATA_MANUAL"]] if "DATA_MANUAL" in idx else None
        data_manual: Optional[date] = None
        if isinstance(data_manual_bruta, datetime):
            data_manual = data_manual_bruta.date()
        elif isinstance(data_manual_bruta, date):
            data_manual = data_manual_bruta
        elif isinstance(data_manual_bruta, str) and data_manual_bruta.strip():
            texto_data = data_manual_bruta.strip()
            data_manual = None
            for formato in ("%d/%m/%Y", "%Y-%m-%d"):  # entrada humana (Pendencias) ou ISO (Resolvidos_Manualmente)
                try:
                    data_manual = datetime.strptime(texto_data, formato).date()
                    break
                except ValueError:
                    continue

        ajustes[source_row] = AjusteManual(
            conta_debito=_texto_ou_none(row[idx["CONTA_DEBITO_MANUAL"]]),
            conta_credito=_texto_ou_none(row[idx["CONTA_CREDITO_MANUAL"]]),
            data=data_manual,
            observacao=_texto_ou_none(row[idx["OBSERVACAO_MANUAL"]]) if "OBSERVACAO_MANUAL" in idx else None,
        )
    return ajustes


# ---------------------------------------------------------------------------
# Aplicação dos ajustes
# ---------------------------------------------------------------------------

@dataclass
class ResultadoAjuste:
    classificacoes: list[Classificacao]
    resolvidos_nesta_execucao: list[Classificacao]  # cópia do estado ANTES de virar OK, para auditoria
    movimentos_corrigidos: dict[int, Movimento]  # cópias com a data corrigida, quando houve DATA_MANUAL


def carregar_source_rows_resolvidos(caminho: str | Path) -> set[int]:
    """Lê só a aba "Resolvidos_Manualmente" e retorna os SOURCE_ROW que já
    foram resolvidos em execuções anteriores — usado para não contar de
    novo, a cada reprocessamento, algo que já tinha sido resolvido antes.
    """
    caminho = Path(caminho)
    if not caminho.exists():
        return set()
    wb = openpyxl.load_workbook(caminho, data_only=True)
    if "Resolvidos_Manualmente" not in wb.sheetnames:
        return set()
    ws = wb["Resolvidos_Manualmente"]
    return {
        int(row[0]) for row in ws.iter_rows(min_row=2, max_col=1, values_only=True) if row[0] is not None
    }


def aplicar_ajustes_manuais(
    classificacoes: list[Classificacao],
    movimentos_por_row: dict[int, Movimento],
    ajustes: dict[int, AjusteManual],
    ja_resolvidos_anteriormente: Optional[set[int]] = None,
) -> ResultadoAjuste:
    """Aplica os ajustes manuais completos sobre as classificações
    pendentes. Nunca reprocessa pela regra automática — o valor manual É
    a decisão final. Itens sem ajuste completo (ou sem ajuste algum)
    saem exatamente como entraram.

    `ja_resolvidos_anteriormente` (opcional) é o conjunto de SOURCE_ROW
    que já constavam na aba "Resolvidos_Manualmente" antes desta
    execução — usado só para que `resolvidos_nesta_execucao` reflita o
    que é NOVO nesta rodada, e não recontar, a cada reprocessamento, uma
    decisão que o conferente já tinha tomado antes.
    """
    ja_resolvidos_anteriormente = ja_resolvidos_anteriormente or set()
    resultado: list[Classificacao] = []
    resolvidos: list[Classificacao] = []
    movimentos_corrigidos: dict[int, Movimento] = {}

    for c in classificacoes:
        if c.status == STATUS_OK:
            resultado.append(c)
            continue

        ajuste = ajustes.get(c.source_row)
        if ajuste is None or not ajuste.contas_completas:
            resultado.append(c)
            continue

        mov = movimentos_por_row.get(c.source_row)
        precisa_data = mov is not None and "data_invalida" in mov.flags
        if precisa_data and ajuste.data is None:
            # contas preenchidas, mas falta a data — continua bloqueado,
            # com o motivo atualizado para deixar claro o que falta.
            pendente = Classificacao(
                c.source_row, c.historico_original, c.chave_encontrada, c.regra_aplicada,
                c.conta_debito, c.conta_credito, c.status,
                detalhes=(c.detalhes + " | Contas preenchidas manualmente, mas falta DATA_MANUAL "
                                        "(a data original é inválida)."),
                flags_fase2=c.flags_fase2,
            )
            resultado.append(pendente)
            continue

        if precisa_data and ajuste.data is not None and mov is not None:
            # a data corrigida precisa chegar até o Movimento em si, não só
            # servir de "senha" para liberar a classificação — senão a
            # Fase 5 (planilha final) não tem de onde pegar a data da
            # linha. Faz uma cópia do Movimento original com a data
            # corrigida; o Movimento de origem (e o .xls) nunca são
            # tocados.
            from dataclasses import replace
            movimentos_corrigidos[c.source_row] = replace(mov, data=ajuste.data)

        resolvido = Classificacao(
            c.source_row, c.historico_original, c.chave_encontrada,
            "ajuste_manual_conferencia",
            ajuste.conta_debito, ajuste.conta_credito, STATUS_OK,
            detalhes=(f"Resolvido manualmente na conferência. Status original: {c.status} "
                      f"({c.regra_aplicada}). Observação do conferente: "
                      f"{ajuste.observacao or '(nenhuma)'}."
                      + (f" | Data corrigida manualmente para {ajuste.data.isoformat()}." if ajuste.data else "")),
            flags_fase2=c.flags_fase2,
        )
        resultado.append(resolvido)
        if c.source_row not in ja_resolvidos_anteriormente:
            resolvidos.append(c)  # só entra na lista de "novidades" se não já registrado antes

    return ResultadoAjuste(classificacoes=resultado, resolvidos_nesta_execucao=resolvidos,
                            movimentos_corrigidos=movimentos_corrigidos)


# ---------------------------------------------------------------------------
# Sugestão de matching (só ajuda visual — nunca aplicada automaticamente)
# ---------------------------------------------------------------------------

def _sugestao_depara(historico_normalizado: str, depara: list[DeParaEntrada]) -> str:
    nomes = sorted({e.nome_original for e in depara if not e.via_alias})
    candidatos = get_close_matches(historico_normalizado, nomes, n=3, cutoff=0.5)
    if not candidatos:
        return ""
    return "Possível(is) correspondência(s), NÃO aplicada(s) automaticamente: " + " | ".join(candidatos)


# ---------------------------------------------------------------------------
# Geração/atualização do arquivo de pendências
# ---------------------------------------------------------------------------

def _motivo(c: Classificacao) -> str:
    base = c.detalhes or ""
    if c.status == STATUS_NAO_ENCONTRADO:
        extra = ("Fornecedor/cliente não encontrado no Controle oficial. Verifique se ele "
                 "precisa ser incluído em a planilha Controle de fornecedores e clientes "
                 "(fora desta planilha de pendências — nunca editamos o Controle automaticamente).")
        return f"{base} {extra}".strip()
    return base


def gerar_ou_atualizar_pendencias(
    classificacoes: list[Classificacao],
    movimentos_por_row: dict[int, Movimento],
    depara: list[DeParaEntrada],
    ano: int,
    mes: int,
    dir_saida: str | Path,
) -> Path:
    """Escreve `output/pendencias_<empresa>_AAAA_MM.xlsx` com os itens
    ainda não-OK. Se o arquivo já existir, preserva qualquer
    preenchimento manual (completo ou parcial) já feito pelo conferente
    para linhas que ainda continuam pendentes — nada se perde entre
    execuções.
    """
    dir_saida = Path(dir_saida)
    dir_saida.mkdir(parents=True, exist_ok=True)
    caminho = dir_saida / nome_arquivo_pendencias(ano, mes)

    ajustes_existentes = carregar_ajustes_manuais(caminho)

    pendentes = [c for c in classificacoes if c.status != STATUS_OK]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pendencias"

    fonte_normal = Font(name=_FONTE, size=10)
    fonte_cabecalho = Font(name=_FONTE, size=10, bold=True, color="FFFFFF")
    preenchimento_cabecalho = PatternFill("solid", fgColor="203864")
    preenchimento_manual = PatternFill("solid", fgColor="FFF2CC")  # amarelo claro: "preencha aqui"

    for col, nome in enumerate(_COLUNAS, start=1):
        cel = ws.cell(row=1, column=col, value=nome)
        cel.font = fonte_cabecalho
        cel.fill = preenchimento_cabecalho
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"

    for i, c in enumerate(pendentes, start=2):
        mov = movimentos_por_row.get(c.source_row)
        ajuste_previo = ajustes_existentes.get(c.source_row)

        valores = {
            "SOURCE_ROW": c.source_row,
            "DATA": mov.data.isoformat() if (mov and mov.data) else None,
            "HISTORICO_ORIGINAL": c.historico_original,
            "DEBITO": mov.debito if mov else None,
            "CREDITO": mov.credito if mov else None,
            "NOME_IDENTIFICADO": c.chave_encontrada or "",
            "STATUS": c.status,
            "MOTIVO": _motivo(c),
            "CONTA_DEBITO_MANUAL": ajuste_previo.conta_debito if ajuste_previo else None,
            "CONTA_CREDITO_MANUAL": ajuste_previo.conta_credito if ajuste_previo else None,
            "DATA_MANUAL": ajuste_previo.data.isoformat() if (ajuste_previo and ajuste_previo.data) else None,
            "OBSERVACAO_MANUAL": ajuste_previo.observacao if ajuste_previo else None,
            "SUGESTAO_DEPARA_AJUDA": (
                _sugestao_depara(mov.historico_normalizado, depara)
                if mov and c.status in (STATUS_NAO_ENCONTRADO, STATUS_AMBIGUO) else ""
            ),
        }
        for col, nome_coluna in enumerate(_COLUNAS, start=1):
            cel = ws.cell(row=i, column=col, value=valores[nome_coluna])
            cel.font = fonte_normal
            if nome_coluna in _COLUNAS_MANUAIS:
                cel.fill = preenchimento_manual

    larguras = {
        "SOURCE_ROW": 10, "DATA": 12, "HISTORICO_ORIGINAL": 46, "DEBITO": 12, "CREDITO": 12,
        "NOME_IDENTIFICADO": 30, "STATUS": 16, "MOTIVO": 50,
        "CONTA_DEBITO_MANUAL": 18, "CONTA_CREDITO_MANUAL": 18, "DATA_MANUAL": 14,
        "OBSERVACAO_MANUAL": 34, "SUGESTAO_DEPARA_AJUDA": 40,
    }
    for col, nome in enumerate(_COLUNAS, start=1):
        ws.column_dimensions[get_column_letter(col)].width = larguras[nome]

    _escrever_aba_instrucoes(wb, ano, mes, total_pendentes=len(pendentes))

    wb.save(caminho)
    return caminho


def _escrever_aba_instrucoes(wb: openpyxl.Workbook, ano: int, mes: int, total_pendentes: int) -> None:
    ws = wb.create_sheet("Instrucoes")
    fonte = Font(name=_FONTE, size=10)
    fonte_titulo = Font(name=_FONTE, size=12, bold=True)

    linhas = [
        (f"Pendências — {cfg.NOME_EMPRESA} — {mes:02d}/{ano}", fonte_titulo),
        ("", fonte),
        (f"Total de movimentos pendentes nesta execução: {total_pendentes}", fonte),
        ("", fonte),
        ("Como usar este arquivo:", fonte_titulo),
        ("1. As colunas até MOTIVO são só leitura — mostram o que o sistema já sabe sobre o movimento.", fonte),
        ("2. Preencha CONTA_DEBITO_MANUAL e CONTA_CREDITO_MANUAL (destacadas em amarelo) com as "
         "contas contábeis corretas para aquele lançamento.", fonte),
        ("3. Se o STATUS for DADO_INVALIDO por causa de data ausente/errada, preencha também "
         "DATA_MANUAL (formato DD/MM/AAAA).", fonte),
        ("4. Use OBSERVACAO_MANUAL para registrar o porquê da decisão — fica gravado na "
         "auditoria do lançamento.", fonte),
        ("5. Uma linha só é considerada resolvida quando AS DUAS contas manuais estiverem "
         "preenchidas (e a data manual também, se for o caso).", fonte),
        ("6. Salve o arquivo com o mesmo nome, no mesmo lugar, e rode o processamento de novo. "
         "O sistema relê exatamente este arquivo.", fonte),
        ("7. A coluna SUGESTAO_DEPARA_AJUDA é só uma pista (comparação de texto, sem "
         "inteligência de negócio) — nunca é aplicada automaticamente. Confira sempre.", fonte),
        ("", fonte),
        ("O que este arquivo NUNCA faz:", fonte_titulo),
        ("- Nunca altera a planilha a planilha Controle de fornecedores e clientes. Se um "
         "fornecedor precisa ser cadastrado, isso é feito à parte, por você, no Controle oficial.", fonte),
        ("- Nunca altera os extratos .xls de origem.", fonte),
        ("- Nunca infere uma conta a partir da outra, nem reaplica regra automática sobre um "
         "valor que você ajustou.", fonte),
    ]
    for i, (texto, f) in enumerate(linhas, start=1):
        cel = ws.cell(row=i, column=1, value=texto)
        cel.font = f
    ws.column_dimensions["A"].width = 100


def escrever_aba_resolvidos(
    caminho: str | Path,
    resolvidos: list[Classificacao],
    ajustes: dict[int, AjusteManual],
) -> None:
    """Acrescenta, numa aba de auditoria permanente, o que foi resolvido
    manualmente — em qualquer execução, não só a atual. Esta aba é a
    memória de longo prazo dos ajustes: diferente da aba "Pendencias"
    (que só lista o que ainda está aberto), o que entra aqui nunca é
    removido, então um reprocessamento futuro nunca "esquece" uma
    decisão já tomada pelo conferente, mesmo depois do item sair da
    lista de pendências.
    """
    if not resolvidos:
        return
    wb = openpyxl.load_workbook(caminho)

    colunas = ["SOURCE_ROW", "HISTORICO_ORIGINAL", "STATUS_ANTERIOR", "REGRA_ANTERIOR",
               "CONTA_DEBITO_MANUAL", "CONTA_CREDITO_MANUAL", "DATA_MANUAL", "OBSERVACAO_MANUAL",
               "RESOLVIDO_EM"]

    if "Resolvidos_Manualmente" in wb.sheetnames:
        ws = wb["Resolvidos_Manualmente"]
        ja_registrados = {
            row[0] for row in ws.iter_rows(min_row=2, max_col=1, values_only=True) if row[0] is not None
        }
        proxima_linha = ws.max_row + 1
    else:
        ws = wb.create_sheet("Resolvidos_Manualmente")
        fonte_cabecalho = Font(name=_FONTE, size=10, bold=True, color="FFFFFF")
        preenchimento_cabecalho = PatternFill("solid", fgColor="375623")
        for col, nome in enumerate(colunas, start=1):
            cel = ws.cell(row=1, column=col, value=nome)
            cel.font = fonte_cabecalho
            cel.fill = preenchimento_cabecalho
        for col, largura in zip("ABCDEFGHI", (10, 46, 18, 22, 18, 18, 14, 34, 20)):
            ws.column_dimensions[col].width = largura
        ja_registrados = set()
        proxima_linha = 2

    agora = datetime.now().strftime("%Y-%m-%d %H:%M")
    fonte_normal = Font(name=_FONTE, size=10)
    for c in resolvidos:
        if c.source_row in ja_registrados:
            continue  # já registrado numa execução anterior — não duplica
        ajuste = ajustes.get(c.source_row, AjusteManual())
        valores = [
            c.source_row, c.historico_original, c.status, c.regra_aplicada,
            ajuste.conta_debito, ajuste.conta_credito,
            ajuste.data.isoformat() if ajuste.data else None,
            ajuste.observacao, agora,
        ]
        for col, valor in enumerate(valores, start=1):
            ws.cell(row=proxima_linha, column=col, value=valor).font = fonte_normal
        proxima_linha += 1

    wb.save(caminho)


# ---------------------------------------------------------------------------
# Regra de bloqueio
# ---------------------------------------------------------------------------

def verificar_bloqueio(classificacoes: list[Classificacao]) -> tuple[bool, dict[str, int]]:
    """Retorna (bloqueado, contagem_por_status)."""
    contagem: dict[str, int] = {}
    for c in classificacoes:
        contagem[c.status] = contagem.get(c.status, 0) + 1
    bloqueado = any(contagem.get(s, 0) > 0 for s in STATUS_BLOQUEANTES)
    return bloqueado, contagem
