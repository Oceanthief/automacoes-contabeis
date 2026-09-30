# Especificação — Automação de Lançamentos CONTAMAX

Status: **auditoria concluída — todas as decisões de negócio respondidas e confirmadas — nenhuma implementação realizada**
Arquivos auditados:
- extrato bancário real de um mês (amostra; não faz parte do repositório)
- TXT de importação do ERP (referência de formato; dados reais não relacionados ao CONTAMAX, não fazem parte do repositório)

---

## 1. Estrutura real do Excel (extrato bancário)

**Biblioteca de leitura**: o arquivo é `.xls` binário (BIFF/OLE2 — "Composite Document File V2"), não `.xlsx`. É preciso usar `xlrd` (que só lê `.xls`), **não** `openpyxl` (só lê `.xlsx`) e **não** `pandas.read_excel` sem especificar o engine `xlrd`.

**Planilha**: 1 única aba, nome `Sheet1`.

**Layout de linhas (0-indexado)**:

| Linha | Conteúdo | Observação |
|---|---|---|
| 0 | Título: `SANTANDER <EMPRESA> PR  - JUNHO  2026` | Célula mesclada A1:E1. Só a coluna A tem texto; B–E vêm como string vazia. Deve ser **ignorada**. |
| 1 | Cabeçalho real: `Data`, `Histórico`, `Crédito`, `Débito`, `Saldo` | Este é o cabeçalho verdadeiro — a leitura de colunas deve mapear por este nome, não por posição fixa "às cegas" (mitiga risco de o banco reordenar colunas em meses futuros). |
| 2 | `SALDO ANTERIOR` | Linha de totalização/abertura, não é movimentação. Crédito e Débito vazios; só `Saldo` preenchido. **Deve ser excluída** do processamento. |
| 3 a 27 | Movimentações reais (25 linhas na amostra) | Cada linha tem Data, Histórico, e **exatamente um** de Crédito/Débito preenchido (o outro vem como célula vazia). |
| última (28 na amostra) | `SALDO FINAL` | Linha de totalização/fechamento, mesma lógica da `SALDO ANTERIOR`. **Deve ser excluída**. |

**Coluna DATA**: tipo de célula `xldate` (ctype 3 no xlrd), serial numérico (ex.: `46174.0`). **Não é string** — não dá pra usar `datetime.strptime` direto. Precisa `xlrd.xldate_as_datetime(valor, wb.datemode)`. `datemode` do workbook = 0 (padrão 1900). Validado: os seriais da amostra convertem para datas de 01/06/2026 a 30/06/2026, batendo com o título "JUNHO 2026".

**Colunas DÉBITO e CRÉDITO**: tipo `float` (ctype 2) quando preenchidas. Quando vazias, **não vêm como `None`** — vêm como ctype 0 com valor `''` (string vazia). O código precisa tratar isso explicitmente (`if cell.ctype == xlrd.XL_CELL_EMPTY or valor == ''`), senão uma soma pode quebrar ou silenciosamente tratar `''` como 0 de forma não intencional.

**Coluna HISTÓRICO**: texto livre, formatos observados: `"Pago: <doc> - <fornecedor>"`, `"TRANSFERENCIA ..."`, `"RECEBIDO NF ..."`, `"TARIFA ..."`, `"JUROS LIMITE CONTA"`, `"IOF"`, `"SALDO ANTERIOR"`, `"SALDO FINAL"`. Não há um padrão fixo de prefixo que sirva de regra — a única forma confiável de excluir as linhas de totalização é comparar o texto exato com `SALDO ANTERIOR` / `SALDO FINAL` (ver seção 5 — Risco).

**Coluna SALDO**: não é usada nos cálculos do CONTAMAX, mas serve como checagem cruzada opcional (saldo da linha anterior ± crédito/débito = saldo da linha atual).

**Linhas com múltiplos lançamentos na mesma data**: confirmado — o agrupamento por data precisa somar vários registros (ex.: 09/06/2026 tem 4 movimentações).

---

## 2. Estrutura real do TXT (referência de formato de importação)

**Importante**: este TXT contém 312 lançamentos reais de **pagamentos a fornecedores** (histórico `PGTO ...`), não tem nenhum exemplo de lançamento CONTAMAX. Ele serve **só** como referência de layout/sintaxe do arquivo que o ERP espera receber.

**Codificação**: UTF-8 com BOM (`EF BB BF` no início do arquivo). Se abrir com `encoding='utf-8'` puro, o primeiro campo da primeira linha vem com o caractere BOM colado (`\ufeff""`). Usar `encoding='utf-8-sig'`.

**Terminador de linha**: `\r\n` (CRLF) em todas as 312 linhas.

**Cabeçalho**: **não existe**. A primeira linha já é dado.

**Delimitador**: vírgula, com cada campo entre aspas duplas: `"campo1","campo2",...,"campo8"`. Delimitador efetivo de parsing é a sequência `","`.

**Quantidade de campos por linha**: **8**, consistente em 100% das 312 linhas (validado).

**Aspas**: não há aspas escapadas (`""` dentro de um campo) nem vírgulas dentro dos campos de texto na amostra — o parsing simples por `","` funciona sem precisar de um parser CSV mais robusto, mas **usar o módulo `csv` do Python é mais seguro** para o caso de um mês futuro trazer um histórico com vírgula.

**Mapeamento dos 8 campos** (posição → conteúdo observado):

| # | Conteúdo observado | Exemplo | Interpretação |
|---|---|---|---|
| 1 | sempre vazio (`""`) | `""` | Campo reservado/não utilizado neste layout |
| 2 | código numérico, **se repete** para o mesmo fornecedor (ex.: um mesmo código aparece 13x para o mesmo transportador) | `"16865"` | Código do favorecido/fornecedor no ERP — **não é sequencial** (não é monotônico ao longo do arquivo, confirmado) |
| 3 | conta contábil, varia entre `11110`, `11212`, `11213` | `"11212"` | Conta de contrapartida do lançamento (não usada por nós — nosso caso usa 11023/11145) |
| 4 | data | `"03/08/2026"` | Formato `DD/MM/AAAA`, string (não é serial Excel aqui, é texto puro) |
| 5 | valor monetário | `"570.86"` | Ponto como separador decimal, sem separador de milhar, sempre 2 casas decimais, sem símbolo de moeda, sempre positivo na amostra |
| 6 | código fixo | `"00013"` | **Constante em 100% das 312 linhas** — confirma a regra do prompt |
| 7 | histórico/descrição | `"PGTO <FORNECEDOR>..."` | Texto livre |
| 8 | referência/documento | `"123456A1B"` | Alfanumérico, **único em cada uma das 312 linhas** (validado) — parece ser número de nota/documento do fornecedor, não um contador sequencial simples |

> **Nota**: este mapeamento (campo 2 = código do fornecedor, campo 3 = conta única) vale para os lançamentos de pagamento a fornecedor que compõem esta amostra. Para os lançamentos CONTAMAX gerados pela automação, os campos 2 e 3 têm semântica diferente — ver seção 5.3.

---

## 3. Regras contábeis confirmadas (definidas pela contabilidade, não inferidas do arquivo)

Para cada data com movimentações:

- **DÉBITO > CRÉDITO**: `VALOR = débito_total - crédito_total`; Conta Débito = `11023`; Conta Crédito = `11145`; Histórico = `VLR REF. RESGATE CONTAMAX`.
- **CRÉDITO > DÉBITO**: `VALOR = crédito_total - débito_total`; Conta Débito = `11145`; Conta Crédito = `11023`; Histórico = `VLR REF. APLICAÇÃO CONTAMAX`.

Simulação feita com os dados reais da amostra (11 datas com movimentação) confirma que a regra é aplicável sem ambiguidade em 100% dos casos da amostra — **nenhuma data teve empate exato** entre débito e crédito, o que significa que o cenário de empate segue sem exemplo real para validar a regra (ver seção 5).

---

## 4. Riscos encontrados

1. **Empate DÉBITO = CRÉDITO em uma data**: não ocorreu na amostra, então não há como inferir do arquivo o que fazer (gerar lançamento com valor 0? não gerar linha para aquela data? outra regra?). Ver seção 5.
2. **Detecção de linhas de totalização por texto exato**: `SALDO ANTERIOR` e `SALDO FINAL` são os únicos textos observados nesta amostra. Se o banco mudar o texto (ex.: "SALDO DO DIA", "SALDO EM 30/06") em outro mês, a exclusão por comparação exata falha silenciosamente e a linha entraria na soma como se fosse movimentação real, distorcendo os totais. Mitigação recomendada: além do texto, também detectar essas linhas por regra estrutural (Crédito **e** Débito ambos vazios, só Saldo preenchido) — isso é mais robusto que comparar string.
3. **Célula vazia do xlrd não é `None`**: é string vazia (`''`) com `ctype` 0. Um código que faça `if debito:` sem tratar isso corretamente pode ter comportamento inesperado dependendo de como a soma é implementada.
4. **Layout do Excel pode variar entre meses/contas**: a amostra é de uma única conta (Santander, filial PR) e um único mês. O título muda a cada mês (nome da conta + mês/ano) — isso não afeta o parsing se ele ignorar a linha 0 por posição, mas é importante não usar o texto do título como fonte de dados (ex.: para inferir o mês de referência), pois ele é apenas informativo e seu formato não é validado por nenhuma regra.
5. **Ordem das colunas do Excel**: a especificação assume que a ordem `Data, Histórico, Crédito, Débito, Saldo` (linha 1) se mantém em arquivos futuros. Recomendação: mapear colunas pelo nome do cabeçalho lido dinamicamente, não pela posição fixa (A, B, C, D, E), para tolerar uma eventual reordenação do banco.
6. **Datas do TXT como string `DD/MM/AAAA`, não como valor Excel**: ao gerar o novo TXT de importação, a data deve sair no mesmo formato string `DD/MM/AAAA`, convertida a partir do date do Excel — não é o mesmo tipo de dado em ambos os arquivos, então há uma conversão obrigatória de tipo (`datetime` → string formatada).
7. **Precisão de ponto flutuante em somas monetárias**: somar dezenas de valores `float` (débito/crédito) pode gerar erros de arredondamento (ex.: `0.1 + 0.2 != 0.3` em ponto flutuante binário). Recomendação para a implementação futura: usar `Decimal` para os totais, não `float` puro.
8. **`xlrd` está com suporte limitado para `.xls` legado**: versões recentes do `xlrd` (≥2.0) removeram suporte a `.xlsx`, mas mantêm suporte a `.xls` — isso é o que precisamos aqui, então não é um bloqueio, mas vale fixar a versão no ambiente para não quebrar com uma atualização futura do pacote.

---

## 5. Decisões pendentes de definição

### 5.1 Empate exato entre total DÉBITO e total CRÉDITO em uma mesma data — RESOLVIDO
**Decisão da contabilidade**: não gerar lançamento para a data (data com débito = crédito é simplesmente ignorada, nenhuma linha sai no TXT para ela).

### 5.2 Geração do campo "SEQUÊNCIA" (campo 8) dos novos lançamentos — RESOLVIDO
**Decisão da contabilidade**: o campo 8 de cada linha gerada recebe um número sequencial simples (1, 2, 3...), na ordem em que as linhas são geradas — não reaproveita nenhum dado do extrato nem do TXT de referência.

**Confirmado pela contabilidade**: a numeração começa em `1` a cada execução do script (não continua de um lote anterior) e não tem zero-padding (`1`, `2`, `3`, ... e não `00001`, `00002`...) — o TXT de referência usa formatos heterogêneos no campo 8 (`"1234"`, `"5678E"`, `"901234..."`), então não há um padrão de largura fixa a seguir mesmo.

### 5.3 Uma linha ou duas linhas por data no TXT gerado — RESOLVIDO
**Decisão da contabilidade**: uma única linha por data. Diferente do TXT de referência (onde campo 2 = código do fornecedor e campo 3 = uma única conta), no lançamento CONTAMAX os campos 2 e 3 carregam as duas contas da partida dobrada lado a lado, na mesma linha: campo 2 = conta débito, campo 3 = conta crédito.

**Exemplo confirmado pela contabilidade** (data 02/09/2026, débito > crédito, diferença = 570.86):
```
"","11023","11145","02/09/2026","570.86","00013","VLR REF. RESGATE CONTAMAX","1"
```

**Layout final da linha CONTAMAX (8 campos, mesma sintaxe do TXT de referência: aspas + vírgula + CRLF + UTF-8 com BOM)**:

| Campo | Conteúdo | Exemplo |
|---|---|---|
| 1 | vazio | `""` |
| 2 | conta débito (11023 ou 11145, conforme a regra da seção 3) | `"11023"` |
| 3 | conta crédito (o outro lado da conta 2) | `"11145"` |
| 4 | data no formato `DD/MM/AAAA` | `"02/09/2026"` |
| 5 | valor da diferença (débito − crédito, ou vice-versa), 2 casas decimais, ponto decimal | `"570.86"` |
| 6 | código fixo | `"00013"` |
| 7 | histórico — `VLR REF. RESGATE CONTAMAX` ou `VLR REF. APLICAÇÃO CONTAMAX` conforme a regra | `"VLR REF. RESGATE CONTAMAX"` |
| 8 | número sequencial simples, incrementado a cada linha gerada (decisão 5.2) | `"1"` |

Isso também resolve a dúvida em aberto da seção 5.2 sobre "1 ou 2 linhas por data": a numeração sequencial do campo 8 conta **uma vez por data processada** (não dobra), já que cada data gera exatamente uma linha.

---

## 6. Esquema dos dados para o processamento (fluxo confirmado)

```
arquivo bruto (.xls)
  → ignorar linha 0 (título) e linha 1 (cabeçalho, usado só para mapear colunas por nome)
  → ignorar linhas cujo Histórico == "SALDO ANTERIOR" ou "SALDO FINAL"
    (reforçar com checagem estrutural: Crédito e Débito ambos vazios)
  → para cada linha restante, extrair:
      - data       (xldate → datetime.date)
      - valor_debito  (float ou 0 se célula vazia)
      - valor_credito (float ou 0 se célula vazia)
  → normalização:
      - data convertida para datetime.date (sem hora)
      - valores numéricos tratados como Decimal (recomendado) para evitar erro de arredondamento
  → agrupamento por data:
      - soma_debito(data)  = Σ valor_debito de todas as linhas daquela data
      - soma_credito(data) = Σ valor_credito de todas as linhas daquela data
  → cálculo CONTAMAX por data:
      - se soma_debito > soma_credito:
          valor = soma_debito - soma_credito
          conta_debito = 11023 | conta_credito = 11145
          historico = "VLR REF. RESGATE CONTAMAX"
      - se soma_credito > soma_debito:
          valor = soma_credito - soma_debito
          conta_debito = 11145 | conta_credito = 11023
          historico = "VLR REF. APLICAÇÃO CONTAMAX"
      - se soma_debito == soma_credito:
          não gerar lançamento para esta data (decisão 5.1)
  → saída esperada (fora de escopo nesta etapa): 1 linha por data processada,
    no layout final confirmado (seção 5.3):
    "","<conta_debito>","<conta_credito>","<data DD/MM/AAAA>","<valor>","00013","<histórico>","<sequência>"\r\n
    codificado em UTF-8 com BOM, campo 6 fixo = "00013", campo 8 = sequencial
    simples incrementado a cada linha gerada (1 por data, decisão 5.2)
```

---

## 7. Resumo — o que falta para liberar a implementação

Todas as decisões de negócio (5.1, 5.2, 5.3, incluindo o critério de início da numeração do campo 8) estão respondidas e confirmadas. A especificação está fechada para iniciar a implementação do script principal.
