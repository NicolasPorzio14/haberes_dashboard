"""
Agrega la hoja `Calendario_Vencimientos` a una planilla YA EXISTENTE (no toca
Movimientos, Supuestos, Obligaciones_Futuras ni Clientes).

A diferencia de `Movimientos` (cosas que YA pasaron), acá se cargan cosas que
TODAVÍA no pasaron pero ya se saben con fecha cierta: vencimientos de
impuestos, sueldos a pagar, cobros esperados de clientes. Es el insumo del
calendario de liquidez de corto plazo del dashboard — no es una proyección
ni un pronóstico, es un registro de compromisos conocidos.

Uso:
    python extend_sheet_calendario.py --credentials credenciales.json \
                                       --spreadsheet-id 14wmeOLLj9TEAuwUFq6f1DwIlasYBub6g05DP1ypdyE4
"""

import argparse

import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import ValidationConditionType

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

CATEGORIAS_COBRO = [
    "Liquidación de Sueldos", "Tercerización de Estudios",
    "Asesoramiento Laboral", "Flujo Sole", "Mandú", "Otro ingreso",
]
CATEGORIAS_GASTO = [
    "Nómina", "Retiros de socios", "Gastos de estructura",
    "Costos externos", "Impuestos IIBB", "Impuestos Ganancias",
    "Cuotas ARCA", "Préstamo Galicia", "Otro gasto",
]

CALENDARIO_HEADERS = ["ID", "Fecha esperada", "Tipo", "Categoría", "Concepto",
                       "Importe", "Estado", "Observación"]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--credentials", required=True)
    parser.add_argument("--spreadsheet-id", required=True)
    args = parser.parse_args()

    creds = Credentials.from_service_account_file(args.credentials, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(args.spreadsheet_id)

    try:
        ws = sh.worksheet("Calendario_Vencimientos")
        print("La hoja 'Calendario_Vencimientos' ya existe — solo se revisan encabezados y validaciones.")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("Calendario_Vencimientos", rows=500, cols=len(CALENDARIO_HEADERS))
        print("Hoja 'Calendario_Vencimientos' creada.")

    ws.update(values=[CALENDARIO_HEADERS], range_name="A1")
    ws.freeze(rows=1)

    ws.add_validation("C2:C500", condition_type=ValidationConditionType.one_of_list,
                       values=["Cobro", "Gasto"], strict=True, showCustomUi=True)
    ws.add_validation("D2:D500", condition_type=ValidationConditionType.one_of_list,
                       values=CATEGORIAS_COBRO + CATEGORIAS_GASTO, strict=False, showCustomUi=True)
    ws.add_validation("G2:G500", condition_type=ValidationConditionType.one_of_list,
                       values=["Pendiente", "Cumplido"], strict=True, showCustomUi=True)

    print("Listo. Cargá ahí los vencimientos y cobros esperados con fecha cierta — "
          "'Tipo' es Cobro o Gasto (igual que en Movimientos), 'Estado' arranca en Pendiente "
          "y lo pasás a Cumplido cuando ya sucedió (y opcionalmente lo cargás también en Movimientos).")


if __name__ == "__main__":
    main()