# Classificação de extrato — como rodar (uso mensal)

Este guia é para quem vai **rodar** o fechamento mensal, não para quem vai
mexer no código. Duas etapas: **preparar** e **finalizar**.

> Para testar sem dados reais, use os arquivos fictícios de `dados_exemplo/`
> (veja o README). Para um cliente real, crie antes `config/configuracoes_locais.py`
> a partir de `config/configuracoes_locais.example.py`.

---

## 1. Onde colocar o extrato mensal

Coloque o arquivo `.xls` do extrato Santander do mês dentro da pasta:

```
input/
```

O arquivo `controle_fornecedores_clientes.xlsx` também mora
em `input/` — ele já está lá e **não muda todo mês** (só quando você
mesmo precisar cadastrar um fornecedor/cliente novo, editando esse
arquivo diretamente no Excel).

---

## 2. Como executar PREPARAR

Abra um terminal na pasta do projeto e rode:

```
python main.py preparar --extrato input/NOME_DO_ARQUIVO.xls
```

O programa identifica sozinho o mês/ano pelo conteúdo do extrato — não
precisa digitar isso.

O que acontece:
- lê e confere o extrato;
- classifica cada lançamento nas contas contábeis;
- separa o que já pode ser lançado do que precisa de conferência humana;
- gera o arquivo de pendências do mês em `output/`.

Duas saídas possíveis:
- **"Nenhuma pendência"** → pode ir direto para a etapa de finalizar.
- **"Existem N pendência(s)"** → o programa **para de propósito**. Vá para
  o passo 3.

---

## 3. Como resolver pendências

Abra o arquivo que o programa indicou, algo como:

```
output/pendencias_<empresa>_2026_08.xlsx
```

Ele tem duas abas:
- **Pendencias** — cada linha é um lançamento que precisa da sua decisão.
  As colunas até `MOTIVO` são só leitura (mostram o que o sistema já
  sabe). As colunas amarelas são para você preencher:
  - `CONTA_DEBITO_MANUAL` e `CONTA_CREDITO_MANUAL` — as duas contas do
    lançamento (**as duas são obrigatórias** para o item sair da lista
    de pendências);
  - `DATA_MANUAL` — só quando o `MOTIVO` falar de data ausente/inválida;
  - `OBSERVACAO_MANUAL` — o porquê da sua decisão (fica gravado na
    auditoria).
- **Instrucoes** — o mesmo passo a passo, dentro do próprio arquivo.

Se o motivo for **"fornecedor/cliente não encontrado"**, confira se ele
precisa ser cadastrado no `controle_fornecedores_clientes.xlsx` (isso é uma edição manual sua, direto no Excel — o programa
nunca mexe nesse arquivo sozinho).

Depois de preencher, **salve o arquivo com o mesmo nome, no mesmo
lugar**, e rode `preparar` de novo (ou já vá direto para `finalizar` —
ele relê as pendências sozinho).

---

## 4. Como executar FINALIZAR

Só funciona quando não sobrar nenhuma pendência. Rode:

```
python main.py finalizar --extrato input/NOME_DO_ARQUIVO.xls
```

O programa:
1. relê o extrato e o Controle;
2. aplica os ajustes que você preencheu no arquivo de pendências;
3. confere que não sobrou nenhuma pendência (se sobrar, ele para e avisa
   — volte ao passo 3);
4. gera a planilha contábil final;
5. gera o TXT de importação para o ERP;
6. confere que as quantidades e os valores batem entre o extrato, a
   planilha final e o TXT.

Se algo estiver errado em qualquer uma dessas checagens, **nada é
gerado** — o programa avisa o motivo e para.

---

## 5. Quais arquivos serão gerados

Tudo em `output/`, nomeado com o ano e o mês (ex.: `2026_08`):

| Arquivo | Quando aparece | Para que serve |
|---|---|---|
| `pendencias_<empresa>_AAAA_MM.xlsx` | na etapa preparar | conferência humana |
| `lancamentos_<empresa>_AAAA_MM.xlsx` | na etapa finalizar | planilha contábil final (layout oficial) |
| `auditoria_<empresa>_AAAA_MM.xlsx` | na etapa finalizar | rastreamento completo até a linha original do extrato |
| `importacao_<empresa>_AAAA_MM.txt` | na etapa finalizar | arquivo para importar no ERP |
| `execucoes.log` | toda vez que o programa roda | histórico das execuções (data/hora, quantidade, arquivos gerados) |

**Nenhum desses arquivos é importado automaticamente no ERP.** A
importação em si continua manual, fora deste programa.

---

## 6. Como interpretar erros

O programa sempre para com uma mensagem explicando o motivo — nunca
continua sozinho depois de um erro. As mensagens mais comuns:

- **`ERRO: extrato não encontrado`** — confira o caminho depois de
  `--extrato`; o arquivo precisa estar em `input/`.
- **`ERRO: Controle não encontrado`** — o arquivo `controle_fornecedores_clientes.xlsx` precisa estar em `input/`.
- **`PROCESSAMENTO NÃO LIBERADO PARA FINALIZAÇÃO`** / **`NÃO
  FINALIZADO: TOTAL_PENDENCIAS = N`** — ainda há pendência(s) no arquivo
  de `output/pendencias_...xlsx`. Volte ao passo 3.
- **`NÃO FINALIZADO (falha ao gerar o TXT)`** — a conferência de ida e
  volta do TXT encontrou uma divergência antes de entregar o arquivo;
  nada é gerado nesse caso. Avise quem mantém o código.
- Qualquer outro erro com um traceback do Python — não é um erro de
  negócio esperado; avise quem mantém o código, com a mensagem
  completa.

---

## Comandos para o uso mensal (resumo)

```
python main.py preparar  --extrato input/NOME_DO_ARQUIVO.xls
#  (preencher output/pendencias_<empresa>_AAAA_MM.xlsx, se houver)
python main.py finalizar --extrato input/NOME_DO_ARQUIVO.xls
```
