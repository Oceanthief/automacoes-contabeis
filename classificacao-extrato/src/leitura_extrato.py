"""
Camada de leitura e normalização do extrato Santander (.xls).

Fase 2 do projeto (segue o contrato de dados levantado com a
contabilidade na Fase 1). Esta camada:

  1. lê um arquivo .xls de extrato mensal;
  2. localiza a tabela de movimentos dentro do arquivo (sem assumir nome
     de aba ou posição fixa de linha/coluna);
  3. separa o que é movimento bancário do que é estrutura (título,
     cabeçalho, saldo anterior, saldo final);
  4. padroniza cada movimento no formato:
         SOURCE_ROW | DATA | HISTORICO_ORIGINAL | DEBITO | CREDITO
  5. valida cada movimento e SINALIZA problemas — nunca descarta uma
     linha problemática silenciosamente.

Fora de escopo desta fase (definido no contrato de dados): De-Para,
contas contábeis, regras de classificação (PAGO/RECEBIDO, transferências,
salários), geração de Excel final ou TXT de importação.

Nenhum arquivo de origem é escrito. Este módulo é somente leitura.
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

import xlrd


# ---------------------------------------------------------------------------
# Normalização de texto (determinística — usada só para localizar a tabela
# e para gerar a chave de comparação do histórico; NÃO é fuzzy matching de
# negócio, isso fica para uma fase futura).
# ---------------------------------------------------------------------------

def _normalizar_texto(valor: Any) -> str:
    """Trim + espaços internos colapsados + maiúsculas + sem acento.

    Usado tanto para localizar cabeçalhos/linhas estruturais quanto para
    gerar a coluna de histórico normalizado. É determinístico: mesma
    entrada sempre gera a mesma saída, sem heurística de similaridade.
    """
    if valor is None:
        return ""
    texto = str(valor).strip()
    texto = " ".join(texto.split())  # colapsa espaços repetidos
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.upper()


# Rótulos de cabeçalho aceitos (já normalizados) para cada coluna esperada.
_ROTULOS_COLUNA = {
    "data": {"DATA"},
    "historico": {"HISTORICO"},
    "credito": {"CREDITO"},
    "debito": {"DEBITO"},
    "saldo": {"SALDO"},
}

_SALDO_ANTERIOR = "SALDO ANTERIOR"
_SALDO_FINAL = "SALDO FINAL"


# ---------------------------------------------------------------------------
# Modelos de dados
# ---------------------------------------------------------------------------

@dataclass
class Movimento:
    """Um lançamento bancário já padronizado, com seus dados de auditoria."""

    source_row: int  # número da linha na planilha original (0-indexado), para rastreio
    data: Optional[date]
    data_bruta: Any  # valor original da célula de data, para auditoria
    historico_original: str
    historico_normalizado: str
    debito: Optional[float]
    credito: Optional[float]
    flags: list[str] = field(default_factory=list)

    @property
    def valido(self) -> bool:
        """Um movimento é considerado válido quando não tem nenhuma flag."""
        return len(self.flags) == 0

    def to_row(self) -> dict:
        """Representação no formato de saída pedido pelo escopo da Fase 2."""
        return {
            "SOURCE_ROW": self.source_row,
            "DATA": self.data.isoformat() if self.data else None,
            "HISTORICO_ORIGINAL": self.historico_original,
            "DEBITO": self.debito,
            "CREDITO": self.credito,
        }


@dataclass
class LinhaEstrutural:
    """Uma linha que não é movimento (título, cabeçalho, saldo, vazia...)."""

    source_row: int
    tipo: str  # "titulo" | "cabecalho" | "saldo_anterior" | "saldo_final" | "vazia" | "outra_aba"
    conteudo: Any = None


@dataclass
class ResumoLeitura:
    arquivo: str
    aba_utilizada: str
    linha_cabecalho: int
    total_linhas_planilha: int
    total_movimentos: int
    total_movimentos_validos: int
    total_movimentos_sinalizados: int
    total_linhas_estruturais: int
    saldo_anterior: Optional[float]
    saldo_final: Optional[float]
    soma_creditos: float
    soma_debitos: float
    fecha_com_saldo: Optional[bool]  # None quando não há saldo anterior/final para conferir


@dataclass
class LeituraResultado:
    movimentos: list[Movimento]
    estrutura: list[LinhaEstrutural]
    resumo: ResumoLeitura

    def movimentos_com_flag(self, flag: Optional[str] = None) -> list[Movimento]:
        if flag is None:
            return [m for m in self.movimentos if m.flags]
        return [m for m in self.movimentos if flag in m.flags]


class ExtratoInvalidoError(Exception):
    """Levantado quando o arquivo não contém uma tabela reconhecível.

    Isso é diferente de "linha problemática" (que é sinalizada, não
    rejeitada) — aqui é o arquivo inteiro que não tem a estrutura mínima
    esperada (ex.: nenhuma aba com cabeçalho Data/Histórico/Crédito/Débito).
    """


# ---------------------------------------------------------------------------
# Localização da tabela dentro do arquivo
# ---------------------------------------------------------------------------

def _localizar_cabecalho(book: xlrd.book.Book) -> tuple[str, int, dict[str, int]]:
    """Procura, em todas as abas do arquivo, a linha de cabeçalho real.

    Não assume nome de aba nem posição fixa — percorre todas as abas e
    todas as linhas iniciais procurando uma linha cujas células batam
    com os rótulos esperados (DATA, HISTORICO, CREDITO, DEBITO). Isso é
    o que permite aceitar "qualquer extrato mensal com a mesma
    estrutura" mesmo que num mês apareça uma aba extra (ex.: o relatório
    de compatibilidade do Excel visto em julho/2026) ou a ordem das
    colunas mude.

    Retorna (nome_da_aba, indice_da_linha_cabecalho, mapa_coluna).
    `mapa_coluna` associa cada rótulo lógico ("data", "historico",
    "credito", "debito", "saldo") ao índice da coluna naquela aba.
    """
    for nome_aba in book.sheet_names():
        sh = book.sheet_by_name(nome_aba)
        limite = min(sh.nrows, 10)  # cabeçalho deve estar nas primeiras linhas
        for r in range(limite):
            valores = [_normalizar_texto(v) for v in sh.row_values(r)]
            mapa: dict[str, int] = {}
            for rotulo, aliases in _ROTULOS_COLUNA.items():
                for idx, v in enumerate(valores):
                    if v in aliases:
                        mapa[rotulo] = idx
                        break
            obrigatorios = {"data", "historico", "credito", "debito"}
            if obrigatorios.issubset(mapa.keys()):
                return nome_aba, r, mapa
    raise ExtratoInvalidoError(
        "Nenhuma aba do arquivo contém uma linha de cabeçalho com as "
        "colunas esperadas (Data, Histórico, Crédito, Débito)."
    )


# ---------------------------------------------------------------------------
# Conversão de valores
# ---------------------------------------------------------------------------

def _converter_data(valor_bruto: Any, ctype: int, datemode: int) -> Optional[date]:
    """Converte uma célula de data do xlrd para `date` do Python.

    Retorna None quando a célula não é uma data válida (célula vazia,
    texto, número fora de faixa etc.) — a chamadora decide o que fazer
    com esse None (nesta camada, sempre vira a flag `data_invalida`,
    nunca um descarte silencioso).
    """
    if ctype != xlrd.XL_CELL_DATE:
        return None
    try:
        ano, mes, dia, _h, _m, _s = xlrd.xldate_as_tuple(valor_bruto, datemode)
        if ano == 0 and mes == 0 and dia == 0:
            return None
        return date(ano, mes, dia)
    except (xlrd.xldate.XLDateError, ValueError):
        return None


def _converter_valor_monetario(valor_bruto: Any, ctype: int) -> Optional[float]:
    """Converte uma célula de Crédito/Débito para float, ou None se vazia.

    Uma célula vazia (`ctype == XL_CELL_EMPTY` ou `XL_CELL_BLANK`, ou
    string vazia) vira `None` — isto é "sem valor", diferente de "valor
    zero" (que é um número 0 explícito, tratado como sinalização própria
    em `_validar`).
    """
    if ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return None
    if ctype == xlrd.XL_CELL_TEXT and str(valor_bruto).strip() == "":
        return None
    if ctype != xlrd.XL_CELL_NUMBER:
        # valor presente mas não numérico (texto, erro, booleano...) —
        # não é um "vazio" comum; será sinalizado como valor inválido
        # pela chamadora através do tipo de retorno especial abaixo.
        raise ValueError(f"célula monetária com tipo inesperado: ctype={ctype}, valor={valor_bruto!r}")
    return round(float(valor_bruto), 2)


# ---------------------------------------------------------------------------
# Validação (sinalização, nunca descarte)
# ---------------------------------------------------------------------------

def _validar(mov: Movimento) -> None:
    """Preenche `mov.flags` in-place. Nunca remove a linha do resultado."""
    if mov.data is None:
        mov.flags.append("data_invalida")

    if not mov.historico_original:
        mov.flags.append("historico_ausente")

    debito_presente = mov.debito is not None
    credito_presente = mov.credito is not None

    if debito_presente and credito_presente:
        mov.flags.append("debito_e_credito_preenchidos")
    elif not debito_presente and not credito_presente:
        mov.flags.append("sem_valor")

    if debito_presente and mov.debito == 0:
        mov.flags.append("valor_zero_debito")
    if credito_presente and mov.credito == 0:
        mov.flags.append("valor_zero_credito")


def _marcar_duplicadas(movimentos: list[Movimento]) -> None:
    """Sinaliza (sem remover) linhas cuja assinatura de negócio se repete.

    Assinatura = (data, histórico normalizado, débito, crédito). Duas
    linhas idênticas nessa assinatura são um indício forte de exportação
    duplicada do banco — mas continuam no resultado, só marcadas, porque
    pode ser coincidência legítima (dois PIX de mesmo valor no mesmo dia,
    por exemplo) e cabe à conferência humana decidir.
    """
    assinaturas: dict[tuple, list[Movimento]] = {}
    for m in movimentos:
        chave = (m.data, m.historico_normalizado, m.debito, m.credito)
        assinaturas.setdefault(chave, []).append(m)

    for grupo in assinaturas.values():
        if len(grupo) > 1:
            linhas_irmas = [m.source_row for m in grupo]
            for m in grupo:
                outras = [r for r in linhas_irmas if r != m.source_row]
                m.flags.append(f"linha_duplicada(irmas={outras})")


# ---------------------------------------------------------------------------
# Função principal
# ---------------------------------------------------------------------------

def ler_extrato(caminho: str | Path) -> LeituraResultado:
    """Lê um extrato Santander (.xls) e retorna movimentos padronizados.

    Parameters
    ----------
    caminho:
        Caminho para o arquivo .xls. O arquivo é aberto em modo somente
        leitura; esta função nunca escreve no arquivo de origem.

    Returns
    -------
    LeituraResultado
        `movimentos`: todas as linhas identificadas como movimento
        bancário, cada uma já com suas `flags` de validação preenchidas
        (lista vazia = linha limpa). Nenhuma linha problemática é
        descartada — ela aparece na lista com a(s) flag(s) que
        explica(m) o problema.
        `estrutura`: linhas que não são movimento (título, cabeçalho,
        saldo anterior, saldo final, linhas vazias), para auditoria.
        `resumo`: contagens e conferência de saldo.
    """
    caminho = str(caminho)
    book = xlrd.open_workbook(caminho, formatting_info=False, on_demand=False)
    nome_aba, linha_cabecalho, mapa_coluna = _localizar_cabecalho(book)
    sh = book.sheet_by_name(nome_aba)

    idx_data = mapa_coluna["data"]
    idx_hist = mapa_coluna["historico"]
    idx_credito = mapa_coluna["credito"]
    idx_debito = mapa_coluna["debito"]

    estrutura: list[LinhaEstrutural] = []
    for r in range(linha_cabecalho + 1):
        estrutura.append(LinhaEstrutural(r, "titulo_ou_cabecalho", sh.row_values(r)))

    linhas_candidatas = list(range(linha_cabecalho + 1, sh.nrows))

    # Localiza a linha de SALDO ANTERIOR (abre o intervalo de movimentos)
    # e a linha de SALDO FINAL (fecha o intervalo) pelo conteúdo do
    # histórico normalizado, sem assumir posição fixa.
    idx_saldo_anterior = None
    idx_saldo_final = None
    saldo_anterior_valor = None
    saldo_final_valor = None

    for r in linhas_candidatas:
        hist_norm = _normalizar_texto(sh.cell_value(r, idx_hist))
        if hist_norm == _SALDO_ANTERIOR and idx_saldo_anterior is None:
            idx_saldo_anterior = r
            cel_saldo = sh.cell(r, mapa_coluna.get("saldo", idx_credito))
            if "saldo" in mapa_coluna:
                v = sh.cell_value(r, mapa_coluna["saldo"])
                saldo_anterior_valor = float(v) if isinstance(v, (int, float)) else None
    for r in reversed(linhas_candidatas):
        hist_norm = _normalizar_texto(sh.cell_value(r, idx_hist))
        if hist_norm == _SALDO_FINAL and idx_saldo_final is None:
            idx_saldo_final = r
            if "saldo" in mapa_coluna:
                v = sh.cell_value(r, mapa_coluna["saldo"])
                saldo_final_valor = float(v) if isinstance(v, (int, float)) else None

    avisos_estrutura = []
    if idx_saldo_anterior is None:
        avisos_estrutura.append("SALDO ANTERIOR não encontrado — todas as linhas após o cabeçalho foram tratadas como possíveis movimentos.")
    if idx_saldo_final is None:
        avisos_estrutura.append("SALDO FINAL não encontrado — todas as linhas até o fim da aba foram tratadas como possíveis movimentos.")

    inicio_movimentos = (idx_saldo_anterior + 1) if idx_saldo_anterior is not None else linhas_candidatas[0]
    fim_movimentos = idx_saldo_final if idx_saldo_final is not None else (linhas_candidatas[-1] + 1)

    if idx_saldo_anterior is not None:
        estrutura.append(LinhaEstrutural(idx_saldo_anterior, "saldo_anterior", saldo_anterior_valor))

    movimentos: list[Movimento] = []
    for r in range(inicio_movimentos, fim_movimentos):
        hist_bruto = sh.cell_value(r, idx_hist)
        hist_original = str(hist_bruto).strip() if hist_bruto is not None else ""
        hist_original = " ".join(hist_original.split())  # remove só espaços supérfluos, preserva o texto
        hist_normalizado = _normalizar_texto(hist_bruto)

        # Linha inteiramente vazia (às vezes sobra uma linha em branco entre
        # o último movimento e o SALDO FINAL) — não é um movimento, mas
        # também não pode ser descartada sem registro.
        linha_valores = sh.row_values(r)
        if all(str(v).strip() == "" for v in linha_valores):
            estrutura.append(LinhaEstrutural(r, "vazia"))
            continue

        data_cell = sh.cell(r, idx_data)
        data_convertida = _converter_data(data_cell.value, data_cell.ctype, book.datemode)

        try:
            debito = _converter_valor_monetario(sh.cell_value(r, idx_debito), sh.cell(r, idx_debito).ctype)
        except ValueError:
            debito = None
        try:
            credito = _converter_valor_monetario(sh.cell_value(r, idx_credito), sh.cell(r, idx_credito).ctype)
        except ValueError:
            credito = None

        mov = Movimento(
            source_row=r,
            data=data_convertida,
            data_bruta=data_cell.value,
            historico_original=hist_original,
            historico_normalizado=hist_normalizado,
            debito=debito,
            credito=credito,
        )
        _validar(mov)
        movimentos.append(mov)

    _marcar_duplicadas(movimentos)

    if idx_saldo_final is not None:
        estrutura.append(LinhaEstrutural(idx_saldo_final, "saldo_final", saldo_final_valor))
    for r in range(fim_movimentos + (1 if idx_saldo_final is not None else 0), sh.nrows):
        estrutura.append(LinhaEstrutural(r, "posterior_ao_saldo_final", sh.row_values(r)))

    soma_creditos = sum(m.credito for m in movimentos if m.credito is not None)
    soma_debitos = sum(m.debito for m in movimentos if m.debito is not None)

    fecha_com_saldo = None
    if saldo_anterior_valor is not None and saldo_final_valor is not None:
        calculado = round(saldo_anterior_valor + soma_creditos - soma_debitos, 2)
        fecha_com_saldo = abs(calculado - round(saldo_final_valor, 2)) < 0.01

    resumo = ResumoLeitura(
        arquivo=caminho,
        aba_utilizada=nome_aba,
        linha_cabecalho=linha_cabecalho,
        total_linhas_planilha=sh.nrows,
        total_movimentos=len(movimentos),
        total_movimentos_validos=len([m for m in movimentos if m.valido]),
        total_movimentos_sinalizados=len([m for m in movimentos if not m.valido]),
        total_linhas_estruturais=len(estrutura),
        saldo_anterior=saldo_anterior_valor,
        saldo_final=saldo_final_valor,
        soma_creditos=round(soma_creditos, 2),
        soma_debitos=round(soma_debitos, 2),
        fecha_com_saldo=fecha_com_saldo,
    )

    resultado = LeituraResultado(movimentos=movimentos, estrutura=estrutura, resumo=resumo)
    resultado.avisos_estrutura = avisos_estrutura  # type: ignore[attr-defined]
    return resultado
