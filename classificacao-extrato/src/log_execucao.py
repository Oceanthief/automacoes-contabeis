"""
Log de execução (Fase 8).

Uma linha por execução, só com o que o enunciado pede: arquivo
processado, data/hora, quantidade de movimentos, OK, pendências, total e
arquivos gerados. Sem stack trace nem caminho interno — isso é para o
terminal na hora do erro, não para o log operacional que um usuário não
técnico pode abrir para conferir o que já rodou.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path


def registrar(
    caminho_log: str | Path,
    comando: str,
    arquivo_processado: str | Path,
    total_movimentos: int,
    qtd_ok: int,
    qtd_pendencias: int,
    arquivos_gerados: list[Path],
) -> None:
    caminho_log = Path(caminho_log)
    caminho_log.parent.mkdir(parents=True, exist_ok=True)

    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    nomes_gerados = ", ".join(Path(a).name for a in arquivos_gerados) if arquivos_gerados else "(nenhum)"
    linha = (
        f"[{agora}] comando={comando} | arquivo={Path(arquivo_processado).name} | "
        f"total={total_movimentos} | OK={qtd_ok} | pendencias={qtd_pendencias} | "
        f"arquivos_gerados={nomes_gerados}\n"
    )
    with open(caminho_log, "a", encoding="utf-8") as f:
        f.write(linha)
