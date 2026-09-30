# Especificação do TXT de importação contábil

**Arquivo de referência analisado:** um TXT real exportado pelo sistema contábil (308 linhas, um mês fechado anterior ao usado no desenvolvimento). O arquivo em si não faz parte do repositório; valores e descrições abaixo foram substituídos por equivalentes fictícios, mantendo o formato.

**Sistema/layout identificado:** não há documentação oficial nem nome de sistema no arquivo. A identificação é por evidência estrutural e de conteúdo: 6 dos 8 códigos de conta distintos usados no TXT (`00032`, `00109`, `00248`, `00276`, `05620`, `20000`) já existiam cadastrados no DE-PARA do projeto, e o layout de 8 colunas com duas colunas vazias nas posições 1 e 6 é exatamente o mesmo já usado desde a Fase 1 do projeto (`SEM INFORMACAO`, `DEBITO`, `CREDITO`, `DATA`, `VALOR`, `SEM INFORMACAO`, `NOVA DESCRICAO`, `SEQUENCIA`). Não é um layout genérico de mercado — é específico deste sistema contábil.

---

## Estrutura física

| Aspecto | Valor |
|---|---|
| ENCODING | UTF-8 com BOM (`\xef\xbb\xbf`) |
| QUEBRA_DE_LINHA | CRLF (`\r\n`) |
| SEPARADOR | vírgula (`,`) |
| DELIMITADOR DE TEXTO | aspas duplas (`"`) em todo campo, sem exceção |
| CABECALHO | não existe — a primeira linha já é dado real (`SEQUENCIA_CONTABIL=1`) |
| RODAPE | não existe — a última linha é dado real, sem linha de totalização |
| Quantidade de campos | 8 em 100% das 308 linhas, sem exceção |
| Linhas especiais | nenhuma encontrada |

---

## Campos

| CAMPO | POSIÇÃO | ORIGEM | TIPO | FORMATO | TAMANHO | OBRIGATÓRIO | REGRA | EXEMPLO | OBSERVAÇÃO |
|---|---|---|---|---|---|---|---|---|---|
| (sem nome) | 1 | — | texto | sempre vazio | 0 | Sim (posição obrigatória, valor sempre vazio) | sempre `""` | `""` | ORIGEM_DESCONHECIDA quanto ao propósito — mas 100% consistente com a coluna `SEM INFORMACAO` já usada desde a Fase 1 deste projeto na mesma posição |
| DEBITO | 2 | `saida_contabil.DEBITO` (pós Fase 6) | texto numérico | 5 dígitos, zero à esquerda | 5 | Sim | nunca convertido para número — preserva zero à esquerda | `"00248"` | Em 308/308 linhas, sempre 5 caracteres |
| CREDITO | 3 | `saida_contabil.CREDITO` (pós Fase 6) | texto numérico | 5 dígitos, zero à esquerda | 5 | Sim | idem DEBITO | `"10001"` | idem |
| DATA | 4 | `saida_contabil.DATA` | texto | `DD/MM/AAAA` | 10 | Sim | dia e mês sempre com 2 dígitos, ano com 4 | `"02/07/2026"` | Nenhuma outra máscara observada nas 308 linhas |
| VALOR | 5 | `saida_contabil.VALOR` | texto numérico | ponto decimal, 2 casas, sempre positivo | variável | Sim | nunca vírgula, nunca separador de milhar, nunca sinal negativo, sempre exatamente 2 casas mesmo quando `.00` | `"1250.00"`, `"35.10"` | Confirma a regra já adotada no projeto (`Decimal`/centavos, valor sempre positivo — pagamento e recebimento já se distinguem por DEBITO/CREDITO, não por sinal) |
| (sem nome) | 6 | — | texto | sempre vazio | 0 | Sim (posição obrigatória, valor sempre vazio) | sempre `""` | `""` | Mesma observação do campo da posição 1 |
| NOVA_DESCRICAO | 7 | `saida_contabil.NOVA_DESCRICAO` | texto | livre, acentuado | 20 a 38 caracteres observados | Sim | sem truncamento visível na amostra | `"VLR REF. RECEB CLIENTES"` | Ver alerta abaixo — conteúdo real observado diverge do que a Fase 5 deste projeto gera para SISPAG e tarifas |
| SEQUENCIA_CONTABIL | 8 | `saida_contabil.SEQUENCIA_CONTABIL` | texto numérico | inteiro, sem zero à esquerda | variável | Sim | contígua de 1 a N dentro do arquivo/competência — reinicia a cada mês | `"1"`, `"308"` | Bate exatamente com a contagem de linhas do arquivo (308 linhas, sequência 1–308) |

---

## Exemplos reconstruídos (Etapa 6)

Reconstrução determinística e literal (byte a byte, incluindo aspas e vírgulas) a partir dos 8 campos, para 5 linhas de perfis diferentes:

| # | Origem (campos) | Linha do TXT reconstruída | Bate com o arquivo real? |
|---|---|---|---|
| 1 (recebimento) | DEBITO=10001, CREDITO=05620, DATA=01/07/2026, VALOR=1250.00, DESC="VLR REF. RECEB CLIENTES", SEQ=1 | `"","10001","05620","01/07/2026","1250.00","","VLR REF. RECEB CLIENTES","1"` | ✅ Sim |
| 4 (tarifa, com typo real no arquivo) | DEBITO=00248, CREDITO=10001, DATA=02/07/2026, VALOR=35.10, DESC="VLR REF. PGO TARIFA BANCARIA", SEQ=4 | `"","00248","10001","02/07/2026","35.10","","VLR REF. PGO TARIFA BANCARIA","4"` | ✅ Sim |
| 5 (tarifa, grafia correta) | DEBITO=00248, CREDITO=10001, DATA=02/07/2026, VALOR=18.50, DESC="VLR REF. PGTO TARIFA BANCARIA", SEQ=5 | `"","00248","10001","02/07/2026","18.50","","VLR REF. PGTO TARIFA BANCARIA","5"` | ✅ Sim |
| 6 (SISPAG granularizado) | DEBITO=00032, CREDITO=10001, DATA=03/07/2026, VALOR=228.00, DESC="VLR REF. PGTO Sispag Fornecedores", SEQ=6 | `"","00032","10001","03/07/2026","228.00","","VLR REF. PGTO Sispag Fornecedores","6"` | ✅ Sim |
| 98 (pagamento comum) | DEBITO=00032, CREDITO=10001, DATA=08/07/2026, VALOR=90.00, DESC="VLR REF. PGTO Sispag Fornecedores", SEQ=98 | `"","00032","10001","08/07/2026","90.00","","VLR REF. PGTO Sispag Fornecedores","98"` | ✅ Sim |

Todas as 5 reconstruções bateram byte a byte com a linha real do arquivo, incluindo aspas, vírgulas e a inconsistência de digitação (`PGO` vs `PGTO`) preservada tal como está no arquivo real.

---

## ⚠️ Alerta importante — conteúdo de NOVA_DESCRICAO diverge do que este projeto gera hoje

Isso é sobre **conteúdo**, não sobre **layout** (o layout está validado). Registro aqui porque afeta diretamente a Fase 7B:

1. **SISPAG granularizado:** no TXT de referência, toda linha de SISPAG usa a descrição genérica `"VLR REF. PGTO Sispag Fornecedores"` — nunca o nome do fornecedor. A Fase 5 deste projeto (`granularizacao/granularizar.py`) gera `"VLR REF. PGTO {NOME}"` (ex.: `"VLR REF. PGTO AGUAS DO VALE SANEAMENTO SA"`). São formatos de conteúdo diferentes para a mesma situação.
2. **Tarifas:** o TXT de referência usa a descrição fixa `"VLR REF. PGTO TARIFA BANCARIA"` (ou `"PGO"`, com erro de digitação, em pelo menos 1 linha). A Fase 3 deste projeto gera `"VLR REF. PGTO " + DESCRICAO original do extrato` (ex.: `"VLR REF. PGTO TAR BLOQUETO ITAU"`).
3. **Recebimentos genéricos:** o TXT de referência usa `"VLR REF. RECEB CLIENTES"` para PIX/boletos recebidos — que é exatamente o `HISTORICO_PADRAO` cadastrado no DE-PARA na Fase 4 para essas regras, mas o `saida_contabil.xlsx` atual carrega em `NOVA_DESCRICAO` o texto gerado na Fase 3 (`"VLR REF. RECEBIMENTO " + descrição original do banco`), não o `HISTORICO_PADRAO` da regra que bateu.

**Decisão necessária antes da Fase 7B:** o TXT final deve usar a `NOVA_DESCRICAO` que já está em `saida_contabil.xlsx` (mais detalhada, rastreável até o fornecedor/lançamento original), ou replicar o padrão mais genérico que o arquivo de referência mostra? Isso muda o que a Fase 7B vai mapear para esse campo.

---

## Comparação com dados do próprio projeto (Etapa 5)

Não há como cruzar o TXT de referência diretamente com `saida_contabil.xlsx` deste projeto porque são competências diferentes — por isso a reconstrução da Etapa 6 usa os próprios campos do TXT como origem. Ainda assim, a comparação de contas contábeis (6/8 códigos já cadastrados no DE-PARA) e a identidade estrutural com o layout de 8 colunas usado desde a Fase 1 são evidência forte de que é o mesmo sistema/layout, não um layout genérico.

---

## STATUS

## **LAYOUT_TXT_VALIDADO**

Todas as 8 posições têm origem e formato comprovados deterministicamente (as 2 posições vazias têm origem semântica desconhecida, mas comportamento 100% determinístico e comprovado: sempre string vazia entre aspas). Nenhum campo obrigatório ficou sem explicação de formato.
