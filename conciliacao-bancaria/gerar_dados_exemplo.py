"""Gera um mês FICTÍCIO completo (agosto/2026) para rodar e testar a conciliação.

Nenhum dado aqui vem de cliente real: empresas, valores, títulos, CNPJs e
contas contábeis são inventados. O objetivo é exercitar TODOS os caminhos do
pipeline com arquivos no mesmo formato dos reais:

  Extrato bancário (layout Itaú, com metadados antes do cabeçalho)
    - recebimentos: cartões, boletos, PIX, TED, rendimentos, estorno de tarifa
    - tarifas bancárias (regra fixa -> conta 00248)
    - SISPAG SALARIOS / TRIBUTOS (DE-PARA exato)
    - SISPAG FORNECEDORES: 1 pagamento (EXATO_1X1) ou vários (EXATO_SOMA)
  Contas Pagas (relatório do sistema de gestão)
    - título com 2 classificações contábeis (juros + parcelamento)
    - pagamentos de outra conta (cartão) -> fora do escopo
    - título com divergência entre VALOR e soma de VALOR_CLASS
  DE-PARA contábil
    - regras EXATO, COMECA_COM e CONTEM, linhas inativas e linha em branco
    - um fornecedor SEM regra -> exige correção manual
  Correções manuais
    - geradas no final, a partir da primeira execução (igual à rotina real)

Uso:
    python gerar_dados_exemplo.py            -> recria dados_exemplo/
"""
from __future__ import annotations

import random
import shutil
import sys
import tempfile
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent
DESTINO = RAIZ / "dados_exemplo"
sys.path.insert(0, str(RAIZ))

from config import CONTA_BANCARIA_ALVO  # noqa: E402

SEMENTE = 2026
CONTA_CARTAO = "341 - Cartão Corporativo"

ARQ_EXTRATO = "extrato_banco_2026_08.xlsx"
ARQ_CONTAS_PAGAS = "contas_pagas_2026_08.xlsx"
ARQ_DE_PARA = "de_para_contabil.xlsx"
ARQ_CORRECOES = "correcoes_manuais.xlsx"

# Fornecedores fictícios -> conta contábil fictícia (None = sem regra no DE-PARA)
FORNECEDORES = {
    "ESCRITORIO CONTABIL MODELO LTDA": "02997",
    "PAPELARIA CENTRAL LTDA": "83860",
    "AGUAS DO VALE SANEAMENTO SA": "00135",
    "ENERGIA REGIONAL DISTRIBUIDORA SA": "00136",
    "LIMPEZA TOTAL SERVICOS LTDA": "50113",
    "TELECOM EXEMPLO SA": "00140",
    "SEGURADORA HORIZONTE SA": "00243",
    "GOVERNO FEDERAL - DIVERSAS": "00410",
    "GRAFICA IMPRESSA LTDA": None,          # sem regra -> correção manual
}
CONTA_CORRECAO_GRAFICA = "00032"


@dataclass
class Pagamento:
    titulo: int
    nome: str
    data: date
    classificacoes: list[tuple[str, str, Decimal]]  # (código, descrição, valor)
    forma: str = "Cobrança Bancária"
    conta: str = CONTA_BANCARIA_ALVO
    valor_informado: Decimal | None = None  # só quando diverge da soma

    @property
    def valor(self) -> Decimal:
        return sum((c[2] for c in self.classificacoes), Decimal("0"))


def _dinheiro(rng: random.Random, minimo: float, maximo: float) -> Decimal:
    return Decimal(str(round(rng.uniform(minimo, maximo), 2)))


def _cnpj_ficticio(n: int) -> str:
    return f"00.000.{n:03d}/0001-00"


def montar_cenario(rng: random.Random):
    """Devolve (movimentos_extrato, pagamentos_contas_pagas)."""
    titulo = iter(range(900101, 999999))
    pagamentos: list[Pagamento] = []

    def pgto(nome, dia, *valores, **kw) -> Pagamento:
        classes = [("2.10.001", "Despesas operacionais", v) for v in valores]
        p = Pagamento(next(titulo), nome, date(2026, 8, dia), classes, **kw)
        pagamentos.append(p)
        return p

    # --- SISPAG FORNECEDORES: grupos de pagamentos que saem num único débito
    grupos_sispag: list[tuple[int, list[Pagamento]]] = []
    plano = [
        (3, ["ESCRITORIO CONTABIL MODELO LTDA"]),
        (5, ["PAPELARIA CENTRAL LTDA", "LIMPEZA TOTAL SERVICOS LTDA"]),
        (7, ["AGUAS DO VALE SANEAMENTO SA"]),
        (10, ["ENERGIA REGIONAL DISTRIBUIDORA SA", "TELECOM EXEMPLO SA", "PAPELARIA CENTRAL LTDA"]),
        (14, ["SEGURADORA HORIZONTE SA"]),
        (18, ["GRAFICA IMPRESSA LTDA", "LIMPEZA TOTAL SERVICOS LTDA"]),
        (21, ["TELECOM EXEMPLO SA"]),
        (26, ["AGUAS DO VALE SANEAMENTO SA", "ESCRITORIO CONTABIL MODELO LTDA"]),
    ]
    for dia, nomes in plano:
        grupos_sispag.append((dia, [pgto(n, dia, _dinheiro(rng, 180, 2400)) for n in nomes]))

    # --- título com 2 classificações (juros + parcelamento), pago via SISPAG
    gov = Pagamento(
        next(titulo), "GOVERNO FEDERAL - DIVERSAS", date(2026, 8, 28),
        [("2.10.053", "Juros s/Parcelamentos", Decimal("124.41")),
         ("2.10.086", "Parcelamento de tributos federais", Decimal("923.99"))],
    )
    pagamentos.append(gov)
    grupos_sispag.append((28, [gov]))

    # --- dois pagamentos ao MESMO fornecedor no mesmo SISPAG: o nome não pode
    # sair duplicado na NOVA DESCRICAO ("X; X" -> "X")
    grupos_sispag.append((31, [pgto("ENERGIA REGIONAL DISTRIBUIDORA SA", 31, Decimal("389.17")),
                               pgto("ENERGIA REGIONAL DISTRIBUIDORA SA", 31, Decimal("212.64"))]))

    # --- pagamentos de OUTRA conta (cartão): ficam fora do escopo da conciliação
    pgto("PAPELARIA CENTRAL LTDA", 12, _dinheiro(rng, 40, 300), forma="Cartão", conta=CONTA_CARTAO)
    divergente = pgto("TELECOM EXEMPLO SA", 19, Decimal("100.00"), Decimal("62.35"),
                      forma="Cartão", conta=CONTA_CARTAO)
    divergente.valor_informado = Decimal("158.71")  # VALOR != soma(VALOR_CLASS) = 162,35

    # --- movimentos do extrato -------------------------------------------------
    mov: list[dict] = []

    def m(dia, texto, valor, razao="", doc=""):
        mov.append({"dia": dia, "texto": texto, "valor": Decimal(valor), "razao": razao, "doc": doc})

    for dia in (3, 4, 5, 6, 7, 10, 11, 12, 13, 14, 17, 18, 19, 20, 21, 24, 25, 26, 27, 28, 31):
        bandeira = rng.choice(["VISA", "MAST", "ELO"])
        m(dia, f"RECEBIMENTO REDE {bandeira} CD{rng.randint(10**9, 10**10 - 1)}",
          _dinheiro(rng, 150, 1500), "REDE ADQUIRENTE EXEMPLO SA", _cnpj_ficticio(1))
    for dia in (4, 11, 18, 25):
        m(dia, "BOLETOS RECEBIDOS", _dinheiro(rng, 900, 6000))
    for i, dia in enumerate((5, 6, 12, 13, 19, 20, 26, 27), start=1):
        m(dia, f"PIX RECEBIDO CLIENTE {i:02d} {dia:02d}/08", _dinheiro(rng, 200, 900),
          f"CLIENTE FICTICIO {i:02d}", f"***.{i:03d}.000-**")
    m(15, "TED RECEBIDA 000.0000.EMPRESA PARCEIRA", _dinheiro(rng, 2000, 5000),
      "EMPRESA PARCEIRA EXEMPLO LTDA", _cnpj_ficticio(2))
    for dia in (7, 14, 21, 28):
        m(dia, "RENDIMENTOS REND PAGO APLIC AUT MAIS", _dinheiro(rng, 3, 40))
    m(22, "BO EST LANC TARIFA", "12.50")

    for dia, texto in [(3, "TAR SISPAG CONCESSION"), (5, "TAR PIX PGTO TRANSF"), (10, "TAR SISPAG TIT OUTRO BCO"),
                       (14, "TAR BLOQUETO ITAU"), (20, "TAR DEPOSITO DIN AGENCIA"), (31, "TAR CTA EMP MENSAL 08/26")]:
        m(dia, texto, -_dinheiro(rng, 1.5, 60))
    m(5, "SISPAG SALARIOS", -_dinheiro(rng, 18000, 24000))
    m(20, "SISPAG SALARIOS", -_dinheiro(rng, 7000, 9000))
    m(20, "SISPAG TRIBUTOS", -_dinheiro(rng, 3000, 5000))
    for dia, grupo in grupos_sispag:
        m(dia, "SISPAG FORNECEDORES", -sum((p.valor for p in grupo), Decimal("0")))

    mov.sort(key=lambda x: x["dia"])
    return mov, pagamentos


# --------------------------------------------------------------- gravação
def gravar_extrato(mov: list[dict], destino: Path) -> None:
    linhas = [
        [None] * 6,
        ["Atualização:", "01/09/2026 09:00:00", None, None, None, None],
        ["Nome:", "EMPRESA EXEMPLO LTDA", None, None, None, None],
        ["Agência:", "0000", None, None, None, None],
        ["Conta:", "00000-0", None, None, None, None],
        [None] * 6,
        ["Lançamentos", None, None, None, None, None],
        ["Periodo:", "01/08/2026 até 31/08/2026", None, None, None, None],
        [None] * 6,
        ["Data", "Lançamento", "Razão Social", "CPF/CNPJ", "Valor (R$)", "Saldo (R$)"],
    ]
    saldo = Decimal("50000.00")
    linhas.append(["31/07/2026", "SALDO ANTERIOR", None, None, None, float(saldo)])
    dia_atual = None
    for x in mov:
        if dia_atual is not None and x["dia"] != dia_atual:
            linhas.append([f"{dia_atual:02d}/08/2026", "SALDO TOTAL DISPONÍVEL DIA", None, None, None, float(saldo)])
        dia_atual = x["dia"]
        saldo += x["valor"]
        linhas.append([f"{x['dia']:02d}/08/2026", x["texto"], x["razao"] or None, x["doc"] or None,
                       float(x["valor"]), None])
    linhas.append([f"{dia_atual:02d}/08/2026", "SALDO TOTAL DISPONÍVEL DIA", None, None, None, float(saldo)])
    pd.DataFrame(linhas).to_excel(destino, index=False, header=False, sheet_name="Lançamentos")


def gravar_contas_pagas(pagamentos: list[Pagamento], destino: Path) -> None:
    linhas = []
    for p in pagamentos:
        valor_titulo = float(p.valor_informado if p.valor_informado is not None else p.valor)
        for i, (cod, desc, v) in enumerate(p.classificacoes):
            data = pd.Timestamp(p.data)
            linhas.append({
                "DESCRICAO": "Matriz", "TITULO": p.titulo, "DATA_VCTO": data,
                "VALOR": valor_titulo if i == 0 else None,
                "VALOR_RECEB": valor_titulo if i == 0 else None, "VALOR_INICIAL": None,
                "NOME": p.nome, "SOLICITANTE": None, "POSSUI_SOLICITANTE": 0,
                "FORMA_PGTO_PREVISTA": p.forma, "COMPETENCIA_ORDEM": 202608,
                "COMPETENCIA": pd.Timestamp(2026, 8, 1), "NF_NUM": f"NF {p.titulo % 10000}",
                "CODUNID": 1, "UNIDADE": 1, "DATA_EFETIVA": data, "DATA_EFETIVA_CH": data,
                "VALOR_CLASS": float(v), "CLASSIFIC_COD": cod, "CLASSIFIC_DESC": desc,
                "HISTORICO": f"Pagamento referente {desc.lower()} - {p.nome.title()}",
                "FORMA_PGTO": p.forma, "CONTA": p.conta, "NUM_CHEQUE": None, "CODUNID_1": 1,
            })
    pd.DataFrame(linhas).to_excel(destino, index=False, sheet_name="Dados")


def gravar_de_para(destino: Path) -> None:
    colunas = ["CHAVE_MATCH", "TIPO_MOVIMENTO", "TIPO_MATCH", "CONTA_CONTABIL", "PRIORIDADE",
               "ATIVO", "HISTORICO_PADRAO", "ORIGEM_TIPO", "OBSERVACAO"]

    def r(chave, tipo, match, conta, hist, prioridade=1, ativo="SIM", obs=""):
        return [chave, tipo, match, conta, prioridade, ativo, hist, "EXEMPLO", obs]

    regras = [r(nome, "PAGAMENTO", "EXATO", conta, f"VLR REF. PGTO {nome}")
              for nome, conta in FORNECEDORES.items() if conta]
    regras += [
        r("SISPAG SALARIOS", "PAGAMENTO", "EXATO", "00301", "VLR REF. PGTO SALARIOS"),
        r("SISPAG TRIBUTOS", "PAGAMENTO", "EXATO", "00410", "VLR REF. PGTO TRIBUTOS"),
        r("BOLETOS RECEBIDOS", "RECEBIMENTO", "EXATO", "05620", "VLR REF. RECEBIMENTO CLIENTES"),
        r("RENDIMENTOS REND PAGO APLIC AUT MAIS", "RECEBIMENTO", "EXATO", "06010",
          "VLR REF. RENDIMENTO APLICACAO"),
        r("BO EST LANC TARIFA", "RECEBIMENTO", "EXATO", "00248", "VLR REF. ESTORNO TARIFA"),
        [None] * len(colunas),  # separador visual: o leitor deve ignorar
        r("RECEBIMENTO REDE", "RECEBIMENTO", "COMECA_COM", "05625", "VLR REF. RECEBIMENTO CARTOES"),
        r("PIX RECEBIDO", "RECEBIMENTO", "COMECA_COM", "05620", "VLR REF. RECEBIMENTO CLIENTES"),
        r("TED RECEBIDA", "RECEBIMENTO", "COMECA_COM", "05620", "VLR REF. RECEBIMENTO CLIENTES"),
        r("EMPRESA PARCEIRA", "RECEBIMENTO", "CONTEM", "05620", "VLR REF. RECEBIMENTO CLIENTES",
          prioridade=2),
        r("TAR ", "PAGAMENTO", "COMECA_COM", "00248", "VLR REF. TARIFA BANCARIA", ativo="NAO",
          obs="Tarifa é regra fixa do motor; linha mantida só como referência."),
        r("<preencher>", "PAGAMENTO", "EXATO", "<preencher>", "", ativo="NAO",
          obs="Modelo para novas regras."),
    ]
    leia_me = pd.DataFrame([
        ["DE-PARA CONTÁBIL — EXEMPLO FICTÍCIO", None],
        ["Objetivo", "Mapear NOVA DESCRICAO para a conta contábil faltante sem adivinhar classificações."],
        ["Conta bancária", "10001"],
        ["Regra de pagamento", "CREDITO permanece 10001; o DE-PARA preenche somente DEBITO."],
        ["Regra de recebimento", "DEBITO permanece 10001; o DE-PARA preenche somente CREDITO."],
        ["Ordem de match", "EXATO → COMECA_COM → CONTEM."],
        ["Ambiguidade", "Duas regras válidas de mesma prioridade => AMBIGUO; não preencher conta."],
        ["Sem regra", "SEM_DE_PARA; não preencher conta."],
        ["Tarifas", "Conta 00248 é regra fixa do motor."],
    ])
    listas = pd.DataFrame({
        "TIPO_MOVIMENTO": ["PAGAMENTO", "RECEBIMENTO", "AMBOS", None],
        "TIPO_MATCH": ["EXATO", "COMECA_COM", "CONTEM", None],
        "ATIVO": ["SIM", "NAO", None, None],
        "STATUS_RESULTADO": ["MATCH_EXATO", "MATCH_REGRA", "AMBIGUO", "SEM_DE_PARA"],
    })
    with pd.ExcelWriter(destino) as w:
        leia_me.to_excel(w, sheet_name="LEIA_ME", index=False, header=False)
        pd.DataFrame(regras, columns=colunas).to_excel(w, sheet_name="DE_PARA", index=False)
        listas.to_excel(w, sheet_name="LISTAS", index=False)


def gravar_correcoes(linhas: list[list], destino: Path) -> None:
    df = pd.DataFrame(linhas, columns=["SEQUENCIA_CONTABIL", "DEBITO_MANUAL", "CREDITO_MANUAL", "OBSERVACAO"])
    df.to_excel(destino, index=False)


# --------------------------------------------------------------- execução
def main() -> None:
    from pipeline import STATUS_TXT_GERADO, processar_mes

    rng = random.Random(SEMENTE)
    mov, pagamentos = montar_cenario(rng)

    if DESTINO.exists():
        shutil.rmtree(DESTINO)
    DESTINO.mkdir()
    gravar_extrato(mov, DESTINO / ARQ_EXTRATO)
    gravar_contas_pagas(pagamentos, DESTINO / ARQ_CONTAS_PAGAS)
    gravar_de_para(DESTINO / ARQ_DE_PARA)

    entradas = [str(DESTINO / a) for a in (ARQ_EXTRATO, ARQ_CONTAS_PAGAS, ARQ_DE_PARA, ARQ_CORRECOES)]

    with tempfile.TemporaryDirectory() as tmp:
        # 1ª execução (exploratória): igual à rotina real, descobre o que ficou SEM_DE_PARA.
        # Uma correção "neutra" é necessária porque o arquivo não pode estar vazio.
        gravar_correcoes([[1, None, None, "placeholder da execução exploratória"]], DESTINO / ARQ_CORRECOES)
        exploratoria = processar_mes(*entradas, str(Path(tmp) / "exploratoria"))
        saida_contabil = pd.read_excel(
            exploratoria.arquivos_gerados["saida_contabil"], dtype={"DEBITO": str, "CREDITO": str}
        )
        pendentes = saida_contabil[saida_contabil["STATUS_DE_PARA"].isin(["SEM_DE_PARA", "AMBIGUO"])]
        correcoes = [
            [int(l.SEQUENCIA_CONTABIL), CONTA_CORRECAO_GRAFICA, "10001", f"VLR REF. PGTO {l.FORNECEDOR}"]
            for l in pendentes.itertuples()
        ]
        if not correcoes:
            raise SystemExit("Cenário inválido: nenhuma pendência para exercitar as correções manuais.")
        gravar_correcoes(correcoes, DESTINO / ARQ_CORRECOES)

        # 2ª execução (final): tem que gerar o TXT validado.
        final = processar_mes(*entradas, str(Path(tmp) / "final"))
        if final.status != STATUS_TXT_GERADO:
            raise SystemExit(f"Cenário inválido: status final {final.status}\n{final.pendencias}\n{final.erros}")

        # Intermediários usados pelos testes unitários de cada fase
        arq = final.arquivos_gerados
        intermediarios = DESTINO / "intermediarios"
        intermediarios.mkdir()
        for chave, nome in [
            ("extrato_padronizado", "extrato_padronizado.xlsx"),
            ("auditoria_sispag", "auditoria_sispag.xlsx"),
            ("saida_enriquecida", "saida_enriquecida.xlsx"),
            ("saida_enriquecida_sispag", "saida_enriquecida_sispag.xlsx"),
            ("relatorio_validacao", "relatorio_validacao.xlsx"),
            ("auditoria_de_para", "auditoria_de_para.xlsx"),
            ("txt", "referencia_importacao_contabil.txt"),
        ]:
            shutil.copy(arq[chave], intermediarios / nome)

    print(f"Dados fictícios gerados em {DESTINO.relative_to(RAIZ)}/")
    print(f"  movimentações no extrato: {final.resumo['movimentacoes_bancarias']}")
    print(f"  SISPAG: {final.resumo['sispag']}")
    print(f"  DE-PARA: {final.resumo['de_para']}")
    print(f"  correções manuais: {len(correcoes)}")
    print(f"  linhas contábeis: {final.resumo['linhas_contabeis']} | total: {final.resumo['total_exportado']}")


if __name__ == "__main__":
    main()
