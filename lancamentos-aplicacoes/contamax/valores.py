"""
Conversão segura de células de planilha para Decimal.

Por que não usar float direto:
- células de Débito/Crédito vêm do xlrd como float64. Somar dezenas de floats
  (ex.: 0.1 + 0.2) pode acumular erro de arredondamento binário que, num
  contexto contábil, é inaceitável mesmo sendo pequeno (um centavo de
  diferença já quebra uma conciliação bancária).
- `Decimal(0.1)` ainda carrega a imprecisão binária original do float
  (`Decimal(0.1) == Decimal('0.1000000000000000055511151231257827021181583404541015625')`).
  Por isso a conversão correta é `Decimal(str(valor))`, não `Decimal(valor)`.
"""
from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

# Réplica local das constantes de ctype do xlrd (0..6). Mantidas como
# inteiros simples (em vez de importar xlrd aqui) para que este módulo não
# dependa da biblioteca de planilhas — ele só entende "tipo de célula" como
# um conceito genérico, o que facilita testar sem precisar montar um
# arquivo .xls de verdade.
CTYPE_EMPTY = 0
CTYPE_TEXT = 1
CTYPE_NUMBER = 2
CTYPE_DATE = 3
CTYPE_BOOLEAN = 4
CTYPE_ERROR = 5
CTYPE_BLANK = 6

DUAS_CASAS = Decimal("0.01")


class ValorInvalidoError(ValueError):
    """Célula de Débito/Crédito com tipo ou conteúdo que não é dinheiro."""


def extrair_valor_monetario(ctype: int, valor: object) -> Decimal:
    """
    Converte uma célula de Débito ou Crédito para Decimal com 2 casas.

    Células vazias (o caso normal: só um de Débito/Crédito é preenchido por
    linha) retornam Decimal('0.00'), não None — isso simplifica a soma
    (dá pra usar `sum(...)` sem checar None em todo lugar) e reflete
    corretamente a semântica contábil ("não houve débito nesta linha").

    Qualquer célula que não seja número nem vazia (texto, erro, booleano,
    ou uma string não numérica) levanta ValorInvalidoError — o chamador
    decide o que fazer (ver contamax.processamento), mas este módulo nunca
    finge que um valor inválido é zero.
    """
    if ctype in (CTYPE_EMPTY, CTYPE_BLANK):
        return Decimal("0.00")

    if ctype == CTYPE_NUMBER:
        try:
            return Decimal(str(valor)).quantize(DUAS_CASAS, rounding=ROUND_HALF_UP)
        except (InvalidOperation, TypeError) as exc:
            raise ValorInvalidoError(
                f"Valor numérico não conversível para Decimal: {valor!r}"
            ) from exc

    # xlrd às vezes representa célula "vazia por fórmula" como string vazia
    # mesmo com ctype TEXT; tratamos isso como vazio genuíno.
    if ctype == CTYPE_TEXT and valor == "":
        return Decimal("0.00")

    raise ValorInvalidoError(
        f"Célula de valor monetário com tipo inesperado (ctype={ctype}, valor={valor!r})"
    )
