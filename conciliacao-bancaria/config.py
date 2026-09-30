"""Configurações centrais da conciliação bancária.

Mantido separado do código de transformação para que mudanças de regra de
negócio (código de conta, nome das colunas de saída) não exijam mexer na
lógica em si.
"""

import os

# Conta bancária (como aparece na coluna CONTA do relatório de Contas Pagas)
# cujos pagamentos entram na conciliação. O valor real de cada cliente fica
# FORA do código: defina a variável de ambiente CONCILIACAO_CONTA_ALVO
# (ver .env.example). O padrão abaixo é o usado pelos dados de exemplo.
CONTA_BANCARIA_ALVO = os.environ.get("CONCILIACAO_CONTA_ALVO", "341 - Conta Movimento")

# Código de conta usado tanto para DEBITO (recebimento) quanto CREDITO (pagamento).
CONTA_PADRAO = "10001"

# Rótulo repetido de propósito: duas colunas de saída se chamam "SEM INFORMACAO".
# Nomes internos (chaves do dicionário) são distintos só para não colidir em
# nenhuma estrutura Python; o rótulo exibido na planilha final é o mesmo.
COLUNAS_SAIDA = [
    ("col_sem_info_1", "SEM INFORMACAO"),
    ("debito", "DEBITO"),
    ("credito", "CREDITO"),
    ("data", "DATA"),
    ("valor", "VALOR"),
    ("col_sem_info_2", "SEM INFORMACAO"),
    ("descricao", "DESCRICAO"),
    ("sequencia", "SEQUENCIA"),
]
