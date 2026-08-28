"""
Agrega la hoja `Clientes` a una planilla YA EXISTENTE (no toca Movimientos
ni Supuestos). Necesaria para Punto de Equilibrio y Análisis de Sensibilidad
"en cantidad de clientes", igual que en el Excel original.

Uso:
    python extend_sheet_clientes.py --credentials credenciales.json \
                                     --spreadsheet-id 14wmeOLLj9TEAuwUFq6f1DwIlasYBub6g05DP1ypdyE4
"""

import argparse

import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

CLIENTES_HEADERS = ["Línea de negocio", "Cantidad de clientes", "Última actualización"]

# Valores iniciales tomados del Excel original (hoja Pto Eq, mayo 2026) —
# reemplazalos por los números actuales del cliente antes de usarlos en serio.
CLIENTES_INICIALES = [
    ["Liquidación de Sueldos", 48, "2026-05-31"],
    ["Tercerización de Estudios", 9, "2026-05-31"],
    ["Asesoramiento Laboral", 1, "2026-05-31"],
    ["Mandú", 6, "2026-05-31"],
]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--credentials", required=True)
    parser.add_argument("--spreadsheet-id", required=True)
    args = parser.parse_args()

    creds = Credentials.from_service_account_file(args.credentials, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(args.spreadsheet_id)

    try:
        ws = sh.worksheet("Clientes")
        print("La hoja 'Clientes' ya existe — no se vuelve a crear, solo se revisan encabezados.")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("Clientes", rows=50, cols=len(CLIENTES_HEADERS))
        print("Hoja 'Clientes' creada.")

    ws.update(values=[CLIENTES_HEADERS] + CLIENTES_INICIALES, range_name="A1")
    ws.freeze(rows=1)
    print("Listo. Revisá los valores de 'Cantidad de clientes' en la planilla — son los de mayo 2026 del Excel original, actualizalos si cambiaron.")


if __name__ == "__main__":
    main()
