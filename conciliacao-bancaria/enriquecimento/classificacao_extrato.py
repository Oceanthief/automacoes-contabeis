"""Enriquecimento do extrato padronizado: coluna NOVA DESCRICAO e classificação
de tarifas bancárias.

Consome a saída da Fase 1 (leitura/extrato_padronizado.py) — não altera
DESCRICAO, DATA, VALOR nem SEQUENCIA. Só decide o texto de NOVA DESCRICAO e,
quando aplicável, reclassifica DEBITO de pagamentos que são tarifa.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from normalizacao import normalizar_conta_contabil, normalizar_texto_classificacao

CONTA_BANCO = "10001"
CONTA_TARIFA = "00248"

DESCRICAO_SISPAG_FORNECEDORES = "SISPAG FORNECEDORES"

PREFIXO_PGTO = "VLR REF. PGTO "
PREFIXO_RECEBIMENTO = "VLR REF. RECEBIMENTO "

# "TAR" ou "TARIFA" como prefixo/palavra completa, não como substring solta
# dentro de outra palavra (ex.: não confundir com "TARDE", "TARGET", etc.).
_REGEX_PREFIXO_TARIFA = re.compile(r"^(TAR|TARIFA)(\s|$)")


class ClassificacaoError(Exception):
    """Erro de classificação: lançamento que não é nem pagamento nem recebimento."""


@dataclass
class Enriquecimento:
    """Resultado da classificação de uma linha do extrato padronizado."""

    nova_descricao: str
    debito: str
    credito: str
    eh_pagamento: bool
    eh_tarifa: bool


def eh_pagamento(debito, credito) -> bool:
    """Pagamento/saída: CREDITO == conta do banco (10001)."""
    return normalizar_conta_contabil(credito) == CONTA_BANCO


def eh_recebimento(debito, credito) -> bool:
    """Recebimento/entrada: DEBITO == conta do banco (10001)."""
    return normalizar_conta_contabil(debito) == CONTA_BANCO


def eh_tarifa_bancaria(descricao: str, debito, credito) -> bool:
    """Tarifa = é pagamento E a descrição normalizada começa com TAR/TARIFA
    como prefixo/palavra (nunca como recebimento, nunca como substring solta).
    """
    if not eh_pagamento(debito, credito):
        return False
    descricao_normalizada = normalizar_texto_classificacao(descricao)
    return bool(_REGEX_PREFIXO_TARIFA.match(descricao_normalizada))


def classificar_linha(descricao: str, debito, credito) -> Enriquecimento:
    """Aplica as Regras 1 e 2 a uma única linha do extrato.

    Levanta ClassificacaoError se a linha não for nem pagamento nem
    recebimento — a Fase 1 garante que isso nunca deveria acontecer, mas o
    enriquecimento não deve silenciar um dado quebrado.
    """
    pagamento = eh_pagamento(debito, credito)
    recebimento = eh_recebimento(debito, credito)

    if not pagamento and not recebimento:
        raise ClassificacaoError(
            f"Lançamento não classificável como pagamento nem recebimento "
            f"(DEBITO={debito!r}, CREDITO={credito!r}, DESCRICAO={descricao!r})"
        )
    if pagamento and recebimento:
        raise ClassificacaoError(
            f"Lançamento com DEBITO e CREDITO simultâneos — dado inválido vindo da "
            f"Fase 1 (DESCRICAO={descricao!r})"
        )

    descricao_str = str(descricao).strip()

    # Exceção: SISPAG FORNECEDORES fica para a próxima fase.
    if descricao_str == DESCRICAO_SISPAG_FORNECEDORES:
        debito_final = normalizar_conta_contabil(debito) or ""
        credito_final = normalizar_conta_contabil(credito) or ""
        return Enriquecimento("", debito_final, credito_final, pagamento, eh_tarifa=False)

    tarifa = eh_tarifa_bancaria(descricao_str, debito, credito)

    if pagamento:
        nova_descricao = PREFIXO_PGTO + descricao_str
        debito_final = CONTA_TARIFA if tarifa else (normalizar_conta_contabil(debito) or "")
        credito_final = CONTA_BANCO
    else:  # recebimento
        nova_descricao = PREFIXO_RECEBIMENTO + descricao_str
        debito_final = CONTA_BANCO
        credito_final = normalizar_conta_contabil(credito) or ""

    return Enriquecimento(nova_descricao, debito_final, credito_final, pagamento, tarifa)
