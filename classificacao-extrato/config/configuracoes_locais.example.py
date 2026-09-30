"""Modelo de configuração local com os dados REAIS do cliente.

Copie para `config/configuracoes_locais.py` (esse nome está no .gitignore)
e redefina só o que for diferente dos valores fictícios de configuracoes.py.
Como as frases abaixo dependem do nome da empresa, redefina-as também.
"""

NOME_EMPRESA = "NOME DA EMPRESA COMO APARECE NO EXTRATO"
PREFIXO_ARQUIVOS = "cliente"

SALARIO_GATILHO_TEXTO = f"SALARIOS- {NOME_EMPRESA}"
TRANSF_ENTRADA_PR_TEXTO = f"TRANSFERENCIA DA CONTA SANTANDER {NOME_EMPRESA} - PR"
TRANSF_SAIDA_PR_TEXTO = f"TRANSFERENCIA PARA CONTA SANTANDER {NOME_EMPRESA} - PR"
TRANSF_SC_TEXTO = f"TRANSFERENCIA PARA CONTA SANTANDER {NOME_EMPRESA} - SC"

ALIASES_DEPARA = {
    # "GRAFIA NO EXTRATO": "NOME NO CONTROLE",
}
