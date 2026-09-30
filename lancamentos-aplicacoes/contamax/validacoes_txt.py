"""
Validação do TXT depois de gravado em disco.

Filosofia igual à de contamax.validacoes (validação pré-Excel): não basta
confiar que o que foi escrito é o que será lido de volta pelo ERP — reler o
arquivo com um parser CSV novo, independente do código que escreveu, é o
jeito de pegar um problema de encoding, de escaping de aspas, ou de
terminador de linha que só aparece no arquivo final, não nos dados em
memória.
"""
from __future__ import annotations

import csv
from datetime import datetime

from .regras import CONTA_APLICACAO_CONTAMAX, CONTA_SALDO_CORRENTE

FORMATO_DATA = "%d/%m/%Y"
CODIGO_FIXO_ESPERADO = "00013"
PARES_CONTA_VALIDOS = {
    (str(CONTA_APLICACAO_CONTAMAX), str(CONTA_SALDO_CORRENTE)),
    (str(CONTA_SALDO_CORRENTE), str(CONTA_APLICACAO_CONTAMAX)),
}


class ErroValidacaoTxt(Exception):
    """Levantada quando o TXT recém-gravado falha em alguma checagem estrutural."""

    def __init__(self, problemas: list[str]):
        self.problemas = problemas
        mensagem = "Validação do TXT encontrou " + str(len(problemas)) + " problema(s):\n- " + "\n- ".join(
            problemas
        )
        super().__init__(mensagem)


def _validar_valor_duas_casas_ponto(valor_str: str) -> bool:
    """Exatamente dígitos.dígitos com 2 casas e ponto decimal — sem regex por preferência de legibilidade."""
    if "," in valor_str:
        return False
    if "." not in valor_str:
        return False
    inteiro, _, decimais = valor_str.partition(".")
    if not inteiro or not inteiro.lstrip("-").isdigit():
        return False
    if len(decimais) != 2 or not decimais.isdigit():
        return False
    return True


def validar_txt_gerado(caminho: str, linhas_esperadas: list[tuple[str, ...]]) -> None:
    """
    Relê `caminho` com csv.reader e confirma, linha a linha:
    1. exatamente 8 campos;
    2. campo 1 vazio;
    3. campo 6 == "00013";
    4. campo 4 (DATA) é uma data válida em DD/MM/AAAA;
    5. campo 5 (VALOR) tem exatamente 2 casas decimais e ponto decimal;
    6. campos 2/3 (DÉBITO/CRÉDITO) formam um dos dois pares de conta válidos;
    7. o conteúdo relido bate exatamente com o que foi gerado em memória
       (linhas_esperadas) — pega qualquer divergência introduzida na escrita.

    Não existe checagem separada de "sem cabeçalho": um cabeçalho de verdade
    (ex. "DATA","CONTA_DEBITO",...) já falharia as checagens de data/valor/
    conta acima, então cai automaticamente nas mesmas regras de qualquer
    outra linha — não precisa de um caso especial.
    """
    problemas: list[str] = []

    with open(caminho, encoding="utf-8-sig", newline="") as arquivo:
        leitor = csv.reader(arquivo, delimiter=",", quotechar='"')
        linhas_lidas = [tuple(linha) for linha in leitor]

    if len(linhas_lidas) != len(linhas_esperadas):
        problemas.append(
            f"Número de linhas lidas ({len(linhas_lidas)}) difere do esperado ({len(linhas_esperadas)})"
        )

    for idx, linha in enumerate(linhas_lidas, start=1):
        if len(linha) != 8:
            problemas.append(f"Linha {idx}: {len(linha)} campos (esperado 8)")
            continue

        campo1, conta_debito, conta_credito, data_str, valor_str, fixo, historico, sequencia = linha

        if campo1 != "":
            problemas.append(f"Linha {idx}: campo 1 não está vazio ({campo1!r})")

        if fixo != CODIGO_FIXO_ESPERADO:
            problemas.append(f"Linha {idx}: campo 6 é {fixo!r}, esperado {CODIGO_FIXO_ESPERADO!r}")

        try:
            datetime.strptime(data_str, FORMATO_DATA)
        except ValueError:
            problemas.append(f"Linha {idx}: data inválida {data_str!r} (esperado DD/MM/AAAA)")

        if not _validar_valor_duas_casas_ponto(valor_str):
            problemas.append(
                f"Linha {idx}: valor {valor_str!r} não tem exatamente 2 casas decimais com ponto"
            )

        if (conta_debito, conta_credito) not in PARES_CONTA_VALIDOS:
            problemas.append(
                f"Linha {idx}: par de contas ({conta_debito}/{conta_credito}) não é um par CONTAMAX válido"
            )

        if not sequencia:
            problemas.append(f"Linha {idx}: campo SEQUÊNCIA vazio")

    if tuple(linhas_lidas) != tuple(linhas_esperadas) and len(linhas_lidas) == len(linhas_esperadas):
        for idx, (lida, esperada) in enumerate(zip(linhas_lidas, linhas_esperadas), start=1):
            if lida != esperada:
                problemas.append(f"Linha {idx}: conteúdo relido {lida!r} difere do gerado {esperada!r}")

    if problemas:
        raise ErroValidacaoTxt(problemas)
