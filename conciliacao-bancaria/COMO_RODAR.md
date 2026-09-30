# Como rodar o processamento mensal — passo a passo

Exemplo usando setembro/2026. Para outro mês, é só trocar os arquivos de entrada.

> **Quer só ver funcionando?** Rode com o mês fictício de `dados_exemplo/`:
>
> ```
> python main.py dados_exemplo/extrato_banco_2026_08.xlsx dados_exemplo/contas_pagas_2026_08.xlsx dados_exemplo/de_para_contabil.xlsx dados_exemplo/correcoes_manuais.xlsx
> ```
>
> Para um cliente real, defina antes a conta bancária em `CONCILIACAO_CONTA_ALVO` (ver `.env.example`).

## 1. Arquivos que você precisa ter em mãos

1. Extrato bancário de setembro (baixado do banco, formato original, sem mexer).
2. Contas Pagas de setembro (exportado do sistema de gestão do cliente).
3. O DE-PARA contábil (`de_para_contabil.xlsx`) — é o mesmo arquivo de sempre. Só edite ele se quiser cadastrar uma regra nova permanente (ver passo 5).
4. Um arquivo de correções manuais (`correcoes_manuais.xlsx`) — para o primeiro mês, comece com um arquivo só com o cabeçalho, sem linhas:

   `SEQUENCIA_CONTABIL | DEBITO_MANUAL | CREDITO_MANUAL | OBSERVACAO`

Não precisam estar em nenhuma pasta especial — você informa o caminho de cada um na hora de rodar.

## 2. Primeira execução (exploratória)

Serve só para descobrir o que a automação não conseguiu classificar sozinha.

```
python main.py "extrato_09_2026.xlsx" "contas_pagas_09_2026.xlsx" "de_para_contabil.xlsx" "correcoes_manuais.xlsx"
```

Essa execução vai processar tudo até a granularização e então parar, avisando que não achou nenhuma correção no arquivo de correções — isso é esperado, porque ele ainda está vazio.

## 3. Ver o que precisa de correção manual

Abra:

```
saida/2026-09/intermediarios/saida_contabil.xlsx
```

Filtre a coluna `STATUS_DE_PARA` por `SEM_DE_PARA` (e também `AMBIGUO`, se aparecer). Para cada linha, anote:

- `SEQUENCIA_CONTABIL`
- `FORNECEDOR`
- `TIPO_MOVIMENTO` (se é PAGAMENTO ou RECEBIMENTO)

## 4. Preencher as correções

No `correcoes_manuais_vazio.xlsx`, adicione uma linha para cada pendência:

- `SEQUENCIA_CONTABIL`: o número que você anotou no passo 3.
- `DEBITO_MANUAL`: preencha se for PAGAMENTO.
- `CREDITO_MANUAL`: preencha se for RECEBIMENTO. Nos dois casos, o outro campo pode ficar `10001` (a conta do banco) ou vazio.
- `OBSERVACAO`: livre, serve só de anotação sua.

Códigos de conta sempre com 5 dígitos, com zero à esquerda quando precisar (ex.: `00248`, não `248`). A única exceção conhecida até hoje é `8685`.

## 5. Se o mesmo fornecedor tende a se repetir todo mês

Em vez de corrigir manualmente de novo no mês que vem, adicione uma regra nova em `de_para_contabil.xlsx` (aba `DE_PARA`), com o nome do fornecedor em `CHAVE_MATCH` e `ATIVO = SIM`. Isso resolve automaticamente dali para frente.

## 6. Apagar a pasta da execução exploratória

```
rm -r saida/2026-09
```

(No Windows: apague a pasta `saida\2026-09` pelo Explorador de Arquivos, ou `rmdir /s saida\2026-09` no PowerShell.)

Isso evita que a execução final vá parar numa pasta `2026-09_v2` sem necessidade.

## 7. Execução final

```
python main.py "extrato_09_2026.xlsx" "contas_pagas_09_2026.xlsx" "de_para_contabil.xlsx" "correcoes_manuais_09_2026.xlsx"
```

## 8. Conferir o resultado

No final da execução, a última linha impressa é o `STATUS`:

- `TXT_GERADO_VALIDADO`: deu tudo certo. O arquivo para importar está em `saida/2026-09/final/2026-09_importacao_contabil.txt`.
- `BLOQUEADO_PARA_TXT`: ainda sobrou alguma pendência. Abra `saida/2026-09/auditorias/relatorio_validacao.xlsx`, aba `ERROS_BLOQUEANTES`, veja o motivo, corrija (normalmente significa que faltou alguma linha em `correcoes_manuais.xlsx`), e repita a partir do passo 6.
- `ERRO_PROCESSAMENTO`: algum arquivo de entrada está com problema. A mensagem de erro impressa no terminal diz qual.

## Onde fica cada coisa, depois de rodar

```
saida/2026-09/
  intermediarios/   arquivos de trabalho (não são o resultado final)
  auditorias/        um Excel de auditoria por fase — é aqui que você investiga qualquer dúvida
  final/              o TXT pronto para importar
  logs/               registro da execução (pipeline.log e resumo_execucao.json)
```
