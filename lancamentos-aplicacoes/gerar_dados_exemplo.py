"""Gera um extrato FICTÍCIO (.xls, agosto/2026) para rodar e testar o CONTAMAX.

A regra do CONTAMAX soma débitos e créditos por dia e gera um único
lançamento com a diferença: RESGATE quando saiu mais do que entrou,
APLICAÇÃO quando entrou mais do que saiu. O cenário cobre:

  - dias só com saídas, só com entradas e mistos (vários lançamentos no dia)
  - um dia de EMPATE exato (débito = crédito): não gera lançamento
  - linhas de SALDO ANTERIOR / SALDO FINAL que não podem entrar na soma

Resultado esperado: 12 dias com movimento -> 11 lançamentos (1 empate).

Uso:
    python gerar_dados_exemplo.py
"""
from __future__ import annotations

from datetime import date
from pathlib import Path

import xlwt

RAIZ = Path(__file__).resolve().parent
DESTINO = RAIZ / "dados_exemplo" / "extrato_2026_08.xls"

# (dia, histórico, crédito, débito)
MOVIMENTOS = [
    (3, "Pago: 1001/08 - FORNECEDOR ALFA LTDA", None, 1250.00),
    (3, "IOF", None, 3.12),
    (4, "RECEBIDO NF 5501 - CLIENTE GAMA LTDA", 4800.00, None),
    (5, "Pago: 1002/08 - TRANSPORTES BETA LTDA", None, 870.45),
    (5, "RECEBIDO NF 5502 - CLIENTE DELTA SA", 300.00, None),
    (7, "TRANSFERENCIA ENTRE CONTAS DA EMPRESA", 2000.00, None),
    (10, "Pago: 1003/08 - FORNECEDOR ALFA LTDA", None, 640.30),
    (10, "Pago: 1004/08 - SERVICOS EPSILON ME", None, 215.00),
    (10, "TARIFA MENSALIDADE PACOTE SERVICOS", None, 49.90),
    (10, "RECEBIDO NF 5503 - CLIENTE GAMA LTDA", 150.00, None),   # 4 movimentos no mesmo dia
    (12, "RECEBIDO NF 5504 - CLIENTE DELTA SA", 1200.00, None),
    (12, "Pago: 1005/08 - TRANSPORTES BETA LTDA", None, 1200.00),  # EMPATE: não gera lançamento
    (14, "JUROS LIMITE CONTA", None, 18.47),
    (17, "RECEBIDO NF 5505 - CLIENTE GAMA LTDA", 3350.75, None),
    (19, "Pago: 1006/08 - FORNECEDOR ALFA LTDA", None, 2210.00),
    (21, "RECEBIDO NF 5506 - CLIENTE DELTA SA", 980.00, None),
    (21, "Pago: 1007/08 - SERVICOS EPSILON ME", None, 409.14),
    (25, "Pago: 1008/08 - TRANSPORTES BETA LTDA", None, 570.86),
    (28, "RECEBIDO NF 5507 - CLIENTE GAMA LTDA", 1890.00, None),
]
SALDO_INICIAL = 500.00


def main() -> None:
    DESTINO.parent.mkdir(exist_ok=True)
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Sheet1")
    f_data = xlwt.easyxf(num_format_str="DD/MM/YYYY")
    f_valor = xlwt.easyxf(num_format_str="#,##0.00")

    ws.write_merge(0, 0, 0, 4, "SANTANDER EMPRESA EXEMPLO PR  - AGOSTO  2026")
    for col, rotulo in enumerate(["Data", "Histórico", "Crédito", "Débito", "Saldo"]):
        ws.write(1, col, rotulo)

    saldo = SALDO_INICIAL
    ws.write(2, 0, date(2026, 7, 31), f_data)
    ws.write(2, 1, "SALDO ANTERIOR")
    ws.write(2, 4, saldo, f_valor)
    linha = 3
    for dia, historico, credito, debito in MOVIMENTOS:
        saldo = round(saldo + (credito or 0) - (debito or 0), 2)
        ws.write(linha, 0, date(2026, 8, dia), f_data)
        ws.write(linha, 1, historico)
        if credito is not None:
            ws.write(linha, 2, credito, f_valor)
        if debito is not None:
            ws.write(linha, 3, debito, f_valor)
        ws.write(linha, 4, saldo, f_valor)
        linha += 1
    ws.write(linha, 0, date(2026, 8, 31), f_data)
    ws.write(linha, 1, "SALDO FINAL")
    ws.write(linha, 4, saldo, f_valor)
    wb.save(str(DESTINO))
    print(f"Extrato fictício gravado em {DESTINO.relative_to(RAIZ)}")


if __name__ == "__main__":
    main()
