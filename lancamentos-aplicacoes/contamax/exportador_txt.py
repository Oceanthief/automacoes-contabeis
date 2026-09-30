"""
Geração do TXT de importação do ERP a partir dos lançamentos já calculados
e já validados (contamax.exportador_xlsx / contamax.validacoes).

Este módulo NÃO recalcula nada — só formata o que o motor contábil (Prompt 2,
não alterado aqui) já produziu, no layout de 8 campos confirmado na
especificação (seção 5.3):

    "","<conta_debito>","<conta_credito>","<data DD/MM/AAAA>","<valor>",
    "00013","<histórico>","<sequência>"

Só dias com status GERADO viram linha no TXT — dias de empate (decisão 5.1)
e dias com erro de valor não geram lançamento, então não têm o que exportar.
"""
from __future__ import annotations

import csv
import os
from dataclasses import dataclass
from decimal import Decimal

from .models import LancamentoContamax, StatusLancamento
from .regras import HISTORICO_APLICACAO, HISTORICO_RESGATE
from .resumo import montar_resumo
from .validacoes_txt import validar_txt_gerado

CODIGO_FIXO = "00013"
FORMATO_DATA = "%d/%m/%Y"


class SequenciaNaoDefinidaError(RuntimeError):
    """
    Levantada quando gerar_txt_contamax é chamado sem uma estratégia de
    sequência explícita.

    A geração do campo 8 (SEQUÊNCIA) foi, na origem, uma decisão de negócio
    pendente (seção 5.2 da especificação) — não algo que dá pra inferir dos
    dados do extrato. Ela já foi respondida pelo usuário, mas este exportador
    continua recusando gerar o TXT sem receber a configuração de forma
    explícita, em vez de assumir silenciosamente um padrão "óbvio". Isso
    protege contra o caso de a decisão nunca ter sido tomada (ou ter sido
    tomada errado) em outro contexto de uso deste mesmo código.
    """


@dataclass(frozen=True)
class EstrategiaSequencia:
    """
    Configuração explícita de como o campo 8 (SEQUÊNCIA) é preenchido.

    Valores confirmados na especificação (seção 5.2): inicio=1,
    zero_padding=None (sem preenchimento com zeros à esquerda), numeração
    reiniciada a cada execução do script.
    """

    inicio: int = 1
    zero_padding: int | None = None  # None = sem padding; N = largura fixa N

    def __post_init__(self) -> None:
        if self.inicio < 1:
            raise ValueError("EstrategiaSequencia.inicio deve ser >= 1")
        if self.zero_padding is not None and self.zero_padding < 1:
            raise ValueError("EstrategiaSequencia.zero_padding deve ser >= 1 quando informado")

    def gerar(self, quantidade: int) -> list[str]:
        valores = range(self.inicio, self.inicio + quantidade)
        if self.zero_padding:
            return [str(v).zfill(self.zero_padding) for v in valores]
        return [str(v) for v in valores]


@dataclass(frozen=True)
class ResumoExportacaoTxt:
    """Números para a mensagem de terminal pedida no Prompt 4."""

    lancamentos: int
    aplicacoes: int
    resgates: int
    erros_estruturais: int


def nome_arquivo_txt_padrao(resultados: list[LancamentoContamax]) -> str:
    """Deriva 'contamax_AAAA-MM.txt' do período dos resultados (não do nome de arquivo de entrada)."""
    resumo = montar_resumo(resultados)
    return f"contamax_{resumo.periodo}.txt"


def _formatar_linha(r: LancamentoContamax, sequencia: str) -> tuple[str, ...]:
    valor_str = str(r.valor.quantize(Decimal("0.01")))
    return (
        "",
        str(r.conta_debito),
        str(r.conta_credito),
        r.data.strftime(FORMATO_DATA),
        valor_str,
        CODIGO_FIXO,
        r.historico,
        sequencia,
    )


def montar_linhas_txt(
    resultados: list[LancamentoContamax], estrategia_sequencia: EstrategiaSequencia
) -> list[tuple[str, ...]]:
    """
    Monta as linhas (já como tuplas de 8 strings) que serão escritas no TXT.
    Separado de gerar_txt_contamax para poder ser testado sem tocar em disco.
    """
    gerados = sorted(
        (r for r in resultados if r.status == StatusLancamento.GERADO), key=lambda r: r.data
    )
    sequencias = estrategia_sequencia.gerar(len(gerados))
    return [_formatar_linha(r, seq) for r, seq in zip(gerados, sequencias)]


def gerar_txt_contamax(
    resultados: list[LancamentoContamax],
    caminho_saida: str,
    estrategia_sequencia: EstrategiaSequencia | None = None,
) -> ResumoExportacaoTxt:
    """
    Escreve o TXT e depois o relê com um parser CSV independente para
    confirmar que o arquivo gravado em disco é estruturalmente válido
    (ver contamax.validacoes_txt) — não basta confiar que o que foi escrito
    é o que seria lido de volta.

    Se a validação pós-escrita falhar, o arquivo inválido é removido do
    disco (não faz sentido deixar um TXT quebrado no lugar de um deliverable).
    """
    if estrategia_sequencia is None:
        raise SequenciaNaoDefinidaError(
            "Nenhuma EstrategiaSequencia foi informada. A geração do campo "
            "SEQUÊNCIA é uma decisão de negócio (especificação, seção 5.2) — "
            "passe explicitamente uma EstrategiaSequencia() para confirmar "
            "que essa decisão foi tomada conscientemente."
        )

    linhas = montar_linhas_txt(resultados, estrategia_sequencia)

    with open(caminho_saida, "w", encoding="utf-8-sig", newline="") as arquivo:
        escritor = csv.writer(
            arquivo, delimiter=",", quotechar='"', quoting=csv.QUOTE_ALL, lineterminator="\r\n"
        )
        escritor.writerows(linhas)

    try:
        validar_txt_gerado(caminho_saida, linhas)
    except Exception:
        if os.path.exists(caminho_saida):
            os.remove(caminho_saida)
        raise

    aplicacoes = sum(1 for linha in linhas if linha[6] == HISTORICO_APLICACAO)
    resgates = sum(1 for linha in linhas if linha[6] == HISTORICO_RESGATE)

    return ResumoExportacaoTxt(
        lancamentos=len(linhas), aplicacoes=aplicacoes, resgates=resgates, erros_estruturais=0
    )
