"""
Configurações centrais da classificação de extrato.

Todo valor de negócio que aparecia espalhado pelos módulos de src/
(contas contábeis fixas, código do ERP, template de histórico, regra de
salário, status da regra de sequência) mora aqui e só aqui. Os módulos em
src/ importam essas constantes; nenhum deles tem esses números soltos no
meio do código de novo.

IMPORTANTE: mudar um valor aqui muda o comportamento contábil real do
sistema. Isso não deve ser feito sem a mesma confirmação formal que cada
regra já recebeu da contabilidade — este arquivo é onde a regra JÁ
CONFIRMADA fica registrada, não um lugar para decidir uma regra nova.

DADOS DO CLIENTE FICAM FORA DO GIT
----------------------------------
Os valores abaixo são de uma EMPRESA FICTÍCIA (usada pelos dados de
exemplo). Para rodar com um cliente real, crie `config/configuracoes_locais.py`
(ignorado pelo Git — ver `configuracoes_locais.example.py`) e redefina lá
só o que muda: nome da empresa, frases do extrato, contas e aliases.
"""

from pathlib import Path

# ---------------------------------------------------------------------------
# Caminhos padrão (relativos à raiz do projeto)
# ---------------------------------------------------------------------------

RAIZ_PROJETO = Path(__file__).resolve().parent.parent
DIR_INPUT = RAIZ_PROJETO / "input"
DIR_OUTPUT = RAIZ_PROJETO / "output"
DIR_DOCS = RAIZ_PROJETO / "docs"

# ---------------------------------------------------------------------------
# Identificação do cliente (valores FICTÍCIOS — sobrescreva localmente)
# ---------------------------------------------------------------------------

NOME_EMPRESA = "EMPRESA EXEMPLO"
PREFIXO_ARQUIVOS = "empresa_exemplo"   # usado no nome dos arquivos de saída

CAMINHO_CONTROLE_PADRAO = DIR_INPUT / "controle_fornecedores_clientes.xlsx"
CAMINHO_LOG_EXECUCOES = DIR_OUTPUT / "execucoes.log"

# ---------------------------------------------------------------------------
# Contas contábeis fixas (confirmadas no contrato de dados, Fase 1)
# ---------------------------------------------------------------------------

CONTA_SANTANDER = "11037"            # conta corrente principal — lado fixo de PAGO/RECEBIDO
CONTA_TRANSFERENCIA_PR = "11023"     # transferência entre contas da empresa - filial PR
CONTA_TRANSFERENCIA_SC = "11221"     # transferência entre contas da empresa - filial SC
CONTA_SALARIO_DIA_06 = "21841"       # débito do salário pago próximo ao dia 6
CONTA_SALARIO_DIA_20 = "11702"       # débito do salário pago próximo ao dia 20
CONTA_PRO_LABORE = "21842"           # débito de pró-labore, independente da data

# ---------------------------------------------------------------------------
# Regra de salário (contrato de dados, Fase 1, Ambiguidade A1 — CONFIRMADA)
# ---------------------------------------------------------------------------

SALARIO_GATILHO_TEXTO = f"SALARIOS- {NOME_EMPRESA}"
SALARIO_ANCORA_1 = 6                 # dia do mês
SALARIO_ANCORA_2 = 20                # dia do mês
SALARIO_TOLERANCIA_DIAS = 3          # ± dias em torno de cada âncora

# ---------------------------------------------------------------------------
# Pró-labore (contrato de dados, Fase 1, Ambiguidade A10 — CONFIRMADA):
# sempre prevalece sobre a regra de salário, mesmo se o histórico contiver
# os dois gatilhos ao mesmo tempo.
# ---------------------------------------------------------------------------

PRO_LABORE_GATILHO_TEXTO = "PRO LABORE"

# ---------------------------------------------------------------------------
# Layout do TXT de importação para o ERP (Fase 6)
# ---------------------------------------------------------------------------

CODIGO_ERP_FIXO = "00013"
TXT_ENCODING = "utf-8-sig"           # grava BOM automaticamente
TXT_DELIMITADOR = ","
TXT_QUOTECHAR = '"'
TXT_QUEBRA_DE_LINHA = "\r\n"
TXT_FORMATO_DATA = "%d/%m/%Y"

# ---------------------------------------------------------------------------
# Novo histórico do lançamento final (Fase 5) — template configurável
# ---------------------------------------------------------------------------

TEMPLATE_NOVO_HISTORICO = "VLR. REF {nome}"

# ---------------------------------------------------------------------------
# Regra da sequência de lançamento (contrato de dados, Fase 1, Ambiguidade
# A3 — CONFIRMADA). Ver lancamentos_finais.py: se este status voltar a
# "PENDENTE_DEFINICAO", a geração da planilha final bloqueia sozinha, sem
# inventar uma regra nova.
# ---------------------------------------------------------------------------

SEQUENCIA_REGRA_STATUS = "CONFIRMADA"
SEQUENCIA_REGRA_DESCRICAO = (
    "Contador simples começando em 1, sem zero-padding, incrementando 1 a 1 "
    "na ordem dos movimentos (SOURCE_ROW), reiniciado a cada execução — "
    "confirmado no contrato de dados da Fase 1 (Ambiguidade A3)."
)
SEQUENCIA_INICIO = 1
SEQUENCIA_ZERO_PADDING = False

# ---------------------------------------------------------------------------
# Marcador literal usado no Controle para contas que exigem preenchimento
# manual (não é uma conta real, nunca deve virar lançamento OK sozinho)
# ---------------------------------------------------------------------------

MARCADOR_PREENCHER_MANUAL = "PREENCHER MANUAL"

# ---------------------------------------------------------------------------
# Transferências entre contas da própria empresa (frases exatas do extrato)
# ---------------------------------------------------------------------------

TRANSF_ENTRADA_PR_TEXTO = f"TRANSFERENCIA DA CONTA SANTANDER {NOME_EMPRESA} - PR"
TRANSF_SAIDA_PR_TEXTO = f"TRANSFERENCIA PARA CONTA SANTANDER {NOME_EMPRESA} - PR"
TRANSF_SC_TEXTO = f"TRANSFERENCIA PARA CONTA SANTANDER {NOME_EMPRESA} - SC"

# ---------------------------------------------------------------------------
# Aliases do DE-PARA: grafias do banco que a contabilidade confirmou serem o
# mesmo cadastro do Controle, mas que ainda não foram incluídas na planilha.
# Ficam aqui (explícitos e auditáveis), nunca escondidos no matching.
#   {grafia_no_extrato: nome_no_controle}
# ---------------------------------------------------------------------------

ALIASES_DEPARA: dict[str, str] = {
    "J P DESPACHOS EXEMPLO": "J.P. DESPACHOS EXEMPLO",
}

# ---------------------------------------------------------------------------
# Sobrescrita local (dados reais do cliente — arquivo fora do Git)
# ---------------------------------------------------------------------------

try:
    from configuracoes_locais import *  # noqa: F401,F403
except ImportError:
    pass
