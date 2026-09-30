# Automações Contábeis em Python

[![testes](https://github.com/Oceanthief/automacoes-contabeis/actions/workflows/testes.yml/badge.svg)](https://github.com/Oceanthief/automacoes-contabeis/actions/workflows/testes.yml)
![Python](https://img.shields.io/badge/Python-3.11+-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-2.x-150458?logo=pandas&logoColor=white)

Três automações que desenvolvi e uso na rotina de fechamento mensal de um escritório de contabilidade. Elas transformam **extratos bancários e relatórios de sistemas de gestão** em **lançamentos contábeis prontos para importar no ERP**, com auditoria em cada etapa.

> **Sobre os dados:** nenhum dado de cliente está neste repositório. Nomes, valores, CNPJs e títulos foram removidos do código. Cada projeto tem um gerador de **dados fictícios** (`gerar_dados_exemplo.py`) que cria um mês completo no mesmo formato dos arquivos reais, e é com ele que os testes e a demonstração rodam. Os dados de cada cliente ficam fora do Git, em configuração local.

## O problema

Todo mês, cada cliente gera centenas de movimentações bancárias que precisam virar lançamentos de partida dobrada (conta débito, conta crédito, histórico). Feito à mão, o trabalho é lento e sujeito a erro:

- um único débito bancário pode pagar vários fornecedores (SISPAG);
- a mesma empresa aparece escrita de formas diferentes no extrato;
- um zero à esquerda perdido num código de conta (`00248` → `248`) faz o ERP rejeitar a importação.

## Os projetos

| Projeto | O que resolve | Destaque técnico | Testes |
|---|---|---|---|
| [**conciliacao-bancaria**](conciliacao-bancaria/) | Concilia o extrato com o relatório de Contas Pagas, classifica cada lançamento e gera o TXT de importação | Conciliação por **soma de subconjuntos** (1 débito = N pagamentos), pipeline de 11 fases com auditoria e trava antes do TXT | 207 |
| [**classificacao-extrato**](classificacao-extrato/) | Classifica o extrato do banco nas contas contábeis a partir de uma planilha de fornecedores e clientes | Motor de regras com **precedência explícita** e ciclo humano no meio (preparar → revisar pendências → finalizar) | 37 |
| [**lancamentos-aplicacoes**](lancamentos-aplicacoes/) | Gera os lançamentos diários de aplicação e resgate automáticos a partir do extrato | Regra de negócio isolada e testável, cálculo em `Decimal`, validação de ida e volta do TXT | 13 |

## Princípios comuns

- **Nunca adivinhar uma conta.** Se nenhuma regra confirmada se aplica, o lançamento vira pendência para revisão humana, e o TXT não é gerado até ela ser resolvida.
- **Nunca alterar o arquivo original.** Toda entrada é somente leitura, e toda saída vai para uma pasta nova.
- **Rastreabilidade.** Cada lançamento final aponta para a linha de origem do extrato, e cada fase grava a própria auditoria em Excel.
- **Dinheiro em `Decimal` ou centavos inteiros**, nunca `float` somado.
- **Validação de ida e volta:** o TXT gerado é relido e comparado com a planilha antes da entrega.
- **Regras de negócio documentadas.** Cada decisão confirmada pela contabilidade está registrada no código ou em `docs/`.

## Como rodar

Cada projeto é independente e tem o próprio `README`. O padrão é o mesmo nos três:

```bash
git clone https://github.com/Oceanthief/automacoes-contabeis.git
cd automacoes-contabeis/conciliacao-bancaria    # ou outro projeto

python -m venv .venv
source .venv/bin/activate                        # Windows: .venv\Scripts\Activate.ps1
pip install -r requirements.txt

pytest                                           # roda a suíte de testes
python gerar_dados_exemplo.py                    # recria o mês fictício
```

## Stack

Python 3.11+ · pandas · openpyxl · xlrd/xlwt (extratos `.xls`) · pytest · GitHub Actions

---

**Victor Ramos Alves de Matos** · Contador (CRC) em transição para dados · [portfólio](https://oceanthief.github.io/victor-ramos-site/)
