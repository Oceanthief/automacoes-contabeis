"""Aplica correcoes_manuais.xlsx sobre a saída contábil da Fase 5.

Opção B combinada com o usuário: nunca editar saida_contabil.xlsx
diretamente. Este módulo devolve uma CÓPIA com as correções aplicadas,
preservando o status original em STATUS_DE_PARA_ORIGINAL e marcando a linha
corrigida como AJUSTE_MANUAL — nunca se confunde com algo que a automação
decidiu sozinha.

Só aceita correção sobre linha que estava genuinamente SEM_DE_PARA ou
AMBIGUO. Tentar "corrigir" uma linha já classificada automaticamente
(MATCH_EXATO, MATCH_REGRA, JA_CLASSIFICADO) é erro bloqueante — nunca
sobrescrita silenciosa.
"""
from __future__ import annotations

import pandas as pd

from leitura.correcoes_manuais import normalizar_codigo_conta_manual

STATUS_AJUSTE_MANUAL = "AJUSTE_MANUAL"
STATUS_ORIGINAIS_CORRIGIVEIS = {"SEM_DE_PARA", "AMBIGUO"}


class AplicarCorrecoesError(Exception):
    """Erro bloqueante ao aplicar uma correção manual."""


def aplicar_correcoes(saida_df: pd.DataFrame, correcoes_df: pd.DataFrame) -> pd.DataFrame:
    """Devolve uma cópia de saida_df com as correções manuais aplicadas."""
    resultado = saida_df.copy()
    resultado["STATUS_DE_PARA_ORIGINAL"] = resultado["STATUS_DE_PARA"]

    sequencias_no_arquivo = set(resultado["SEQUENCIA_CONTABIL"])

    for _, correcao in correcoes_df.iterrows():
        seq = int(correcao["SEQUENCIA_CONTABIL"])
        if seq not in sequencias_no_arquivo:
            raise AplicarCorrecoesError(
                f"correcoes_manuais.xlsx referencia SEQUENCIA_CONTABIL {seq}, "
                "que não existe em saida_contabil.xlsx."
            )

        indice = resultado.index[resultado["SEQUENCIA_CONTABIL"] == seq][0]
        linha = resultado.loc[indice]

        status_atual = str(linha["STATUS_DE_PARA"]).strip()
        if status_atual not in STATUS_ORIGINAIS_CORRIGIVEIS:
            raise AplicarCorrecoesError(
                f"SEQUENCIA_CONTABIL {seq}: tentativa de correção manual sobre linha que já "
                f"tinha STATUS_DE_PARA='{status_atual}' (só SEM_DE_PARA/AMBIGUO podem ser corrigidas)."
            )

        debito_manual = normalizar_codigo_conta_manual(correcao.get("DEBITO_MANUAL"))
        credito_manual = normalizar_codigo_conta_manual(correcao.get("CREDITO_MANUAL"))
        tipo_movimento = str(linha["TIPO_MOVIMENTO"]).strip()

        if tipo_movimento == "PAGAMENTO":
            if not debito_manual:
                raise AplicarCorrecoesError(
                    f"SEQUENCIA_CONTABIL {seq}: é um PAGAMENTO, precisa de DEBITO_MANUAL preenchido."
                )
            credito_existente = str(linha["CREDITO"]).strip()
            if credito_manual and credito_manual != credito_existente:
                raise AplicarCorrecoesError(
                    f"SEQUENCIA_CONTABIL {seq}: CREDITO_MANUAL ({credito_manual}) diverge do "
                    f"CREDITO já existente ({credito_existente}) — pagamento SISPAG deveria manter 10001."
                )
            resultado.at[indice, "DEBITO"] = debito_manual

        elif tipo_movimento == "RECEBIMENTO":
            if not credito_manual:
                raise AplicarCorrecoesError(
                    f"SEQUENCIA_CONTABIL {seq}: é um RECEBIMENTO, precisa de CREDITO_MANUAL preenchido."
                )
            debito_existente = str(linha["DEBITO"]).strip()
            if debito_manual and debito_manual != debito_existente:
                raise AplicarCorrecoesError(
                    f"SEQUENCIA_CONTABIL {seq}: DEBITO_MANUAL ({debito_manual}) diverge do "
                    f"DEBITO já existente ({debito_existente}) — recebimento deveria manter 10001."
                )
            resultado.at[indice, "CREDITO"] = credito_manual

        else:
            raise AplicarCorrecoesError(
                f"SEQUENCIA_CONTABIL {seq}: TIPO_MOVIMENTO inesperado ('{tipo_movimento}')."
            )

        resultado.at[indice, "STATUS_DE_PARA"] = STATUS_AJUSTE_MANUAL

    return resultado
