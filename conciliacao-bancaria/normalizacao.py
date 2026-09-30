"""Regras de normalização de valores monetários, datas e texto de classificação.

Isolado num módulo próprio porque a Fase 0 já identificou que Valor (R$) e
Saldo (R$) precisam de tratamento explícito de tipo — nunca inferência
automática do pandas. A normalização de texto para classificação (maiúsculas,
sem acento) segue o mesmo princípio: nunca decidir algo "no olho", sempre via
regra explícita e testável.
"""
import re
import unicodedata
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal

import pandas as pd

# Única chave de configuração para a regra de sinal do valor exportado.
# True  -> exporta sempre positivo (valor absoluto) — regra desta versão.
# False -> preserva o sinal original da movimentação (pagamento negativo).
# Mudar aqui é suficiente; nenhuma outra parte do código precisa saber disso.
EXPORTAR_VALOR_ABSOLUTO = True


def valor_para_decimal(valor: float) -> Decimal:
    """Converte float para Decimal com 2 casas decimais.

    Usa Decimal(str(valor)) em vez de Decimal(valor) para não herdar o erro
    de representação binária do float (ex.: Decimal(0.1) != Decimal("0.1")).
    """
    return Decimal(str(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def valor_exportacao(valor: float) -> Decimal:
    """Aplica a regra de sinal definida em EXPORTAR_VALOR_ABSOLUTO."""
    decimal_valor = valor_para_decimal(valor)
    return abs(decimal_valor) if EXPORTAR_VALOR_ABSOLUTO else decimal_valor


def para_centavos(valor: Decimal) -> int:
    """Converte um Decimal (2 casas) para inteiro de centavos.

    Usado pelo motor de conciliação para fazer subset-sum em inteiros —
    nunca em float, que acumula erro de arredondamento em somas repetidas.
    """
    return int((valor * 100).to_integral_value(rounding=ROUND_HALF_UP))


def data_para_datetime(data: str) -> pd.Timestamp:
    """Converte string dd/mm/aaaa do extrato para Timestamp.

    format= explícito de propósito: sem isso, o pandas pode interpretar
    01/02/2026 como 2 de janeiro em vez de 1 de fevereiro dependendo da
    versão/locale, um bug silencioso e perigoso em dados financeiros.
    """
    return pd.to_datetime(data, format="%d/%m/%Y")


def normalizar_data(valor) -> date:
    """Normaliza qualquer representação de data para um `date` puro (sem hora).

    Aceita, nesta ordem de tentativa:
      - pandas.Timestamp / datetime.datetime -> descarta a hora;
      - datetime.date -> devolve como está;
      - texto 'dd/mm/aaaa' ou 'aaaa-mm-dd' -> parse explícito;
      - texto em outro formato -> último recurso via pandas (dayfirst=True);
      - número -> tratado como serial de data do Excel (dias desde 1899-12-30).

    Levanta ValueError com uma mensagem clara em vez de deixar o pandas
    inferir algo errado silenciosamente (mesmo cuidado do resto do projeto:
    dados financeiros não podem depender de inferência automática).
    """
    if valor is None:
        raise ValueError("Data ausente — não é possível normalizar um valor None.")

    try:
        if pd.isna(valor):
            raise ValueError("Data ausente — não é possível normalizar um valor nulo (NaN/NaT).")
    except TypeError:
        pass  # pd.isna não aceita todo tipo; segue a checagem normal abaixo

    if isinstance(valor, pd.Timestamp):
        return valor.date()

    if isinstance(valor, datetime):
        return valor.date()

    if isinstance(valor, date):
        return valor

    if isinstance(valor, str):
        texto = valor.strip()
        for formato in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(texto, formato).date()
            except ValueError:
                continue
        try:
            return pd.to_datetime(texto, dayfirst=True, errors="raise").date()
        except Exception as exc:
            raise ValueError(f"Não foi possível interpretar a data '{valor}': {exc}") from exc

    if isinstance(valor, (int, float)):
        # Serial de data do Excel: dias desde 1899-12-30 (convenção do próprio Excel).
        return pd.to_datetime(valor, unit="D", origin="1899-12-30").date()

    raise ValueError(f"Tipo de data não suportado: {type(valor)} ({valor!r})")


def normalizar_texto_classificacao(texto: str) -> str:
    """Normaliza um texto apenas para fins de CLASSIFICAÇÃO (nunca para exibição).

    Regras: maiúsculas, sem acento, sem espaço nas pontas, espaços múltiplos
    viram um só. O texto original passado pelo chamador nunca é alterado —
    esta função sempre devolve uma string nova.
    """
    sem_acento = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode("ascii")
    maiusculo = sem_acento.upper().strip()
    return re.sub(r"\s+", " ", maiusculo)


def normalizar_conta_contabil(valor) -> str:
    """Normaliza um valor de conta contábil vindo do Excel para comparação.

    O Excel/pandas costuma devolver contas numéricas como float ao reler um
    arquivo (ex.: "10001" volta como 10001.0) e células vazias como NaN.
    Esta função devolve sempre uma string comparável: "" para vazio/nulo,
    "10001" para 10001.0 ou "10001", etc. — sem alterar o valor original em
    nenhuma outra parte do código.
    """
    if valor is None:
        return ""
    if isinstance(valor, float):
        if pd.isna(valor):
            return ""
        if valor.is_integer():
            return str(int(valor))
        return str(valor)
    return str(valor).strip()
