# CONTAMAX — automação mensal

Gera o Excel de conferência e o TXT de importação do ERP a partir do
extrato bancário do mês.

## 1. Onde colocar o Excel

Copie o extrato `.xls` do mês (o arquivo que o internet banking do
Santander exporta) para dentro desta pasta — a mesma pasta onde está o
`main.py`. Pode dar qualquer nome ao arquivo (ex.: `extrato_do_mes.xls`).

## 2. Comando para executar

```
python main.py "extrato_do_mes.xls"
```

Troque `"extrato_do_mes.xls"` pelo nome do arquivo que você copiou.
É só isso — nenhum outro parâmetro é necessário. O período (mês/ano) é
detectado automaticamente a partir das datas dentro do próprio extrato.

## 3. Onde encontrar os resultados

Os dois arquivos saem em uma pasta nova, criada automaticamente, com o
padrão:

```
output/AAAA-MM/contamax_AAAA-MM.xlsx   ← Excel de conferência
output/AAAA-MM/contamax_AAAA-MM.txt    ← TXT para importar no ERP
```

Por exemplo, para um extrato de junho de 2026:

```
output/2026-06/contamax_2026-06.xlsx
output/2026-06/contamax_2026-06.txt
```

Se você rodar o programa duas vezes para o mesmo mês, ele **não** vai
sobrescrever o que já existe — vai avisar e parar, pra você não perder um
resultado por engano. Se isso acontecer e você quiser gerar de novo, apague
(ou mova para outro lugar) os dois arquivos daquele mês antes de rodar de
novo.

## 4. Como interpretar erros

Se algo der errado, o programa para e mostra uma mensagem em português
explicando o que aconteceu, embaixo do título `NÃO FOI POSSÍVEL CONCLUIR`.
Os casos mais comuns:

- **"Não encontrei o arquivo..."** — o caminho digitado está errado, ou o
  arquivo não está na pasta que você indicou.
- **"...extensão '.xlsx', que esta automação ainda não sabe ler"** — o
  arquivo precisa estar no formato antigo do Excel (`.xls`). Se o seu
  arquivo é `.xlsx`, abra no Excel e "Salvar como" → `.xls`.
- **"Não encontrei as colunas esperadas..."** — o arquivo não parece ser um
  extrato bancário no formato de sempre (Data, Histórico, Crédito, Débito,
  Saldo). Confira se é o arquivo certo.
- **"Já existe(m) resultado(s) gerado(s) para o período..."** — você já
  rodou para esse mês antes. Veja o item 3 acima.
- **"Encontrei problema(s) que impedem gerar o Excel/TXT com segurança"** —
  uma conferência interna encontrou algo que não bate (por exemplo, um
  valor que não deveria ser negativo). Leia a lista de problemas listada
  logo abaixo dessa mensagem — ela descreve exatamente o que foi encontrado.

Se aparecer `ALGO INESPERADO ACONTECEU`, é um erro que o programa não sabia
prever — anote a mensagem técnica mostrada e leve para quem cuida desta
automação.

Quando um dia específico do mês tem um problema (por exemplo, um valor que
não pôde ser lido, ou um dia em que débito e crédito ficaram exatamente
iguais), o programa **não trava** por causa disso — ele processa o resto do
mês normalmente e avisa no resumo final quantos dias ficaram como
"pendência (revisar manualmente)". Veja esses dias na aba
`AUDITORIA_DIARIA` do Excel, coluna `STATUS`.

## 5. Como conferir o Excel

Abra `contamax_AAAA-MM.xlsx`. Ele tem duas abas:

- **LANCAMENTOS** — o que vai ser importado no ERP: data, conta débito,
  conta crédito, histórico e valor. Uma linha por dia.
- **AUDITORIA_DIARIA** — o "porquê" de cada linha de LANCAMENTOS: o total
  de débito e de crédito daquele dia, a diferença, qual regra foi aplicada
  (resgate ou aplicação) e o status. Todo dia do mês aparece aqui, inclusive
  os que **não** geraram lançamento (empate entre débito e crédito, ou erro
  de leitura) — é aqui que você confere isso antes de importar qualquer
  coisa no ERP.

## 6. Quando importar o TXT

Só importe o `contamax_AAAA-MM.txt` no ERP **depois** de conferir o Excel
(item 5) e concluir que os lançamentos de LANCAMENTOS estão certos. O TXT
já foi conferido automaticamente pelo programa (estrutura, datas, valores),
mas essa conferência automática garante que o arquivo está bem formado —
não substitui você olhar se os *números* fazem sentido para o mês.
