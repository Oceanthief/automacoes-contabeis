# Lançamentos diários de aplicação e resgate automáticos

A conta corrente do cliente tem **aplicação automática** (CONTAMAX): o banco aplica o saldo que sobra e resgata o que falta. No extrato, isso não aparece como lançamento próprio. É preciso deduzir, dia a dia, se houve aplicação ou resgate, e de quanto. Esta automação lê o extrato `.xls` e gera o Excel de conferência e o TXT de importação do ERP.

## A regra

Para cada dia com movimentação, somam-se as saídas (débito) e as entradas (crédito):

| Situação no dia | Lançamento | Débito | Crédito | Histórico |
|---|---|---|---|---|
| Saiu mais do que entrou | **Resgate** da diferença | 11023 (aplicação) | 11145 (conta corrente) | `VLR REF. RESGATE CONTAMAX` |
| Entrou mais do que saiu | **Aplicação** da diferença | 11145 | 11023 | `VLR REF. APLICAÇÃO CONTAMAX` |
| Débito = crédito | nenhum lançamento | — | — | — |

Todas as decisões de negócio (empate, numeração da sequência, uma linha por dia) e os riscos levantados antes da implementação estão em [`docs/especificacao_contamax.md`](docs/especificacao_contamax.md).

## Arquitetura

```
contamax/
├── leitura_xls.py       # acha o cabeçalho pelo nome, trata célula vazia do xlrd
├── valores.py           # converte para Decimal com segurança
├── regras.py            # a regra contábil: pura, sem I/O, testável isolada
├── processamento.py     # leitura + agrupamento + regra
├── validacoes.py        # conferências antes de gravar qualquer coisa
├── exportador_xlsx.py   # Excel de conferência (lançamentos + auditoria diária)
├── exportador_txt.py    # TXT no layout do ERP (UTF-8 com BOM, CRLF)
└── validacoes_txt.py    # relê o TXT gerado e confere com o Excel
```

A regra contábil fica isolada de leitura e escrita de arquivos, por isso dá para testá-la com três números, sem montar um Excel.

## Rodando com o mês fictício

```bash
pip install -r requirements.txt
python main.py dados_exemplo/extrato_2026_08.xls
```

```
Dias processados: 12
Total de resgates: 6
Total de aplicações: 5
Pendências (empate ou valor inválido — revisar manualmente): 1
TXT para importar no ERP: output/2026-08/contamax_2026-08.txt (11 lançamento(s))
```

O extrato de exemplo ([`gerar_dados_exemplo.py`](gerar_dados_exemplo.py)) inclui um dia com 4 movimentações e **um dia de empate**, caso que nunca apareceu nos dados reais e só existia como decisão documentada. Se rodar duas vezes para o mesmo mês, o programa recusa em vez de sobrescrever.

Passo a passo de uso mensal: [`COMO_RODAR.md`](COMO_RODAR.md).

## Testes

```bash
pytest          # 13 testes
```

Cobrem a regra (resgate, aplicação, empate), a soma em `Decimal` (`0,10 + 0,20 = 0,30` exato), o layout do TXT conferido contra o exemplo validado pela contabilidade, e as mensagens de erro para quem não é técnico.
