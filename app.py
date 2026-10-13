"""
Dashboard financiero en tiempo real: Estudio Haberes.

Lee el libro diario (hoja "Movimientos") y los parámetros (hojas "Supuestos",
"Obligaciones_Futuras", "Clientes" y "Calendario_Vencimientos") de una planilla
de Google Sheets, vía la API de Google con una cuenta de servicio.

Carga de datos: SOLO directa en Google Sheets (no hay formulario de carga
dentro de esta app, por decisión explícita del cliente/consultor).

Correr con:
    streamlit run app.py

Fuentes de datos, en este orden de prioridad:
  1. Google Sheets: secrets `gcp_service_account` y `spreadsheet_id`
     (ver SETUP_GOOGLE_CLOUD.md).
  2. Excel local: variable de entorno HABERES_XLSX con la ruta a una copia de
     la planilla de carga (sirve para probar cambios sin tocar la planilla real).
  3. MODO DEMO con datos sintéticos, si no hay ninguna de las dos.

Cambios de la versión 2 (octubre 2026), pensados para trabajar con varios meses:
  * Los Cobros de la categoría "Préstamo Galicia" son FINANCIACIÓN: ya no se
    suman a los ingresos de gestión, ni a los KPIs, ni a la participación por
    línea, ni a la sensibilidad. Lo mismo la amortización de capital (Gastos
    de "Préstamo Galicia"). Ambos se muestran en una sección de financiación y
    sí entran en el flujo de caja.
  * Los KPIs del mes salen del Estado de Resultados (antes mezclaban resultado
    contable con caja: sumaban retiros, cuotas ARCA y préstamo como egresos).
  * El Estado de Resultados muestra cada categoría por separado (Nómina,
    Costos externos, Otros gastos, Estructura, IIBB, Ganancias) y un puente
    explícito entre resultado de gestión y flujo de caja.
  * "Otro ingreso" (Cloud Recibo Digital, etc.) pasa a ser una línea más en la
    rentabilidad por línea, así las participaciones suman 100%.
  * Sección nueva de evolución mensual con variaciones contra el mes anterior.
  * Control de calidad de la carga: avisa categorías desconocidas, combinaciones
    Tipo/Categoría inválidas y partidas provisorias (PARCIAL, PENDIENTE,
    ESTIMADO, A CONFIRMAR) del mes que se está mirando.
  * Corrección: el decorador de caché estaba sobre la función que parsea
    fechas y no sobre la que lee Movimientos.
"""

import datetime as dt
import os
import re

import gspread
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.utils import ValueRenderOption

# ──────────────────────────────────────────────────────────────────────────
# Configuración y categorías (deben coincidir con setup_sheets.py y con las
# listas desplegables de la hoja Movimientos)
# ──────────────────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Ingresos que forman parte del resultado de gestión
CATEGORIAS_INGRESO_GESTION = [
    "Liquidación de Sueldos", "Tercerización de Estudios",
    "Asesoramiento Laboral", "Flujo Sole", "Mandú", "Otro ingreso",
]
# Entradas de plata que NO son ingreso (financiación)
CATEGORIAS_FINANCIACION = ["Préstamo Galicia"]

CATEGORIAS_COBRO = CATEGORIAS_INGRESO_GESTION + CATEGORIAS_FINANCIACION
CATEGORIAS_GASTO = [
    "Nómina", "Retiros de socios", "Gastos de estructura",
    "Costos externos", "Impuestos IIBB", "Impuestos Ganancias",
    "Cuotas ARCA", "Préstamo Galicia", "Otro gasto",
]

# Rubro de cada Gasto. Mismo criterio que la hoja `ER` del Excel original:
# Cuotas ARCA, Retiros de socios y la amortización del préstamo NO son costo
# de gestión; impactan la caja pero no el resultado.
RUBRO_GASTO = {
    "Nómina": "nomina",
    "Costos externos": "costos_externos",
    "Otro gasto": "otros_gastos",          # incluye intereses del préstamo
    "Gastos de estructura": "estructura",
    "Impuestos IIBB": "iibb",
    "Impuestos Ganancias": "ganancias",
    "Cuotas ARCA": "arca",
    "Retiros de socios": "retiros",
    "Préstamo Galicia": "financiacion_salida",
}

RUBROS = [
    "ingresos", "financiacion_entrada", "nomina", "costos_externos", "otros_gastos",
    "estructura", "iibb", "ganancias", "arca", "retiros", "financiacion_salida",
]

# Líneas de negocio para la rentabilidad por línea. "Otro ingreso" se incluye
# para que las participaciones sumen 100% del ingreso de gestión.
LINEAS_DE_NEGOCIO = [
    "Liquidación de Sueldos", "Tercerización de Estudios", "Asesoramiento Laboral",
    "Flujo Sole", "Mandú", "Otro ingreso",
]

# Palabras que marcan una partida como provisoria (en Cliente/Proveedor u
# Observación). Se usan solo para avisar, no cambian ningún cálculo.
MARCAS_PROVISORIAS = re.compile(r"PARCIAL|PENDIENTE|ESTIMAD|A CONFIRMAR", re.IGNORECASE)

CACHE_TTL_SEGUNDOS = 60

# Paleta y tipografía, pensadas para el mundo de un estudio contable (libro
# diario, sellos, papel). Los mismos colores se usan en el CSS y en Plotly.
INK = "#1F2D4A"        # tinta / texto principal
PAPER = "#F6F4EE"      # fondo, papel
CARD = "#FFFFFF"
VERDE_LEDGER = "#3B7A57"   # cumple objetivo
ROJO_SELLO = "#B23A48"     # no cumple objetivo
DORADO_SELLO = "#C08A28"   # acento / advertencia
GRIS_TEXTO = "#6B6558"     # texto secundario

st.set_page_config(page_title="Estudio Haberes | Dashboard", layout="wide")

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:wght@600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@500;600&display=swap');

html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
    color: {INK};
}}
h1, h2, h3 {{
    font-family: 'Source Serif 4', serif !important;
    color: {INK} !important;
    border-bottom: 1px solid #DAD5C7;
    padding-bottom: 0.3rem;
}}
/* Números en monoespaciada, como columnas de un libro contable */
[data-testid="stMetricValue"] {{
    font-family: 'IBM Plex Mono', monospace;
    color: {INK};
}}
[data-testid="stMetric"] {{
    background-color: {CARD};
    border: 1px solid #E4DFD1;
    border-top: 3px solid {INK};
    border-radius: 6px;
    padding: 0.9rem 1rem 0.6rem 1rem;
}}
[data-testid="stDataFrame"] {{
    font-family: 'IBM Plex Mono', monospace;
}}
.semaforo-badge {{
    display: inline-block;
    font-family: 'Inter', sans-serif;
    font-weight: 600;
    font-size: 0.85rem;
    padding: 0.15rem 0.7rem;
    border-radius: 999px;
    margin-right: 0.5rem;
}}
.semaforo-ok {{
    background-color: {VERDE_LEDGER}22;
    color: {VERDE_LEDGER};
    border: 1px solid {VERDE_LEDGER}55;
}}
.semaforo-no {{
    background-color: {ROJO_SELLO}22;
    color: {ROJO_SELLO};
    border: 1px solid {ROJO_SELLO}55;
}}
.semaforo-aviso {{
    background-color: {DORADO_SELLO}22;
    color: {DORADO_SELLO};
    border: 1px solid {DORADO_SELLO}55;
}}
</style>
""", unsafe_allow_html=True)


def plotly_layout_base(fig: go.Figure, **kwargs) -> go.Figure:
    """Aplica la misma identidad visual (tipografía, colores, fondo
    transparente para integrarse con el tema) a todos los gráficos."""
    fig.update_layout(
        font=dict(family="Inter, sans-serif", color=INK),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10, r=10, t=30, b=10),
        **kwargs,
    )
    return fig


def pesos(v) -> str:
    """Formato argentino: punto de miles y signo adelante ($1.234.567 / −$1.234)."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return ""
    if v != v:  # NaN
        return ""
    txt = f"{abs(v):,.0f}".replace(",", ".")
    return f"−${txt}" if round(v) < 0 else f"${txt}"


def porcentaje(v, signo: bool = False) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return ""
    if v != v:
        return "nuevo"
    txt = (f"{v:+.1%}" if signo else f"{v:.1%}").replace(".", ",")
    return txt


# ──────────────────────────────────────────────────────────────────────────
# Fuente de datos
# ──────────────────────────────────────────────────────────────────────────


def _hay_secrets_sheets() -> bool:
    try:
        return "gcp_service_account" in st.secrets and "spreadsheet_id" in st.secrets
    except Exception:  # no existe secrets.toml
        return False


def fuente_de_datos() -> str:
    if _hay_secrets_sheets():
        return "sheets"
    ruta = os.environ.get("HABERES_XLSX", "")
    if ruta and os.path.exists(ruta):
        return "excel"
    return "demo"


def modo_demo_activo() -> bool:
    return fuente_de_datos() == "demo"


@st.cache_resource(show_spinner=False)
def get_client() -> gspread.Client:
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )
    return gspread.authorize(creds)


def _leer_hoja(sheet_name: str) -> pd.DataFrame:
    """Devuelve la hoja como DataFrame, venga de Google Sheets o del Excel
    local. Si la hoja no existe, levanta KeyError."""
    if fuente_de_datos() == "excel":
        try:
            return pd.read_excel(os.environ["HABERES_XLSX"], sheet_name=sheet_name,
                                 keep_default_na=False, na_values=[""])
        except ValueError as e:  # hoja inexistente
            raise KeyError(sheet_name) from e
    client = get_client()
    sh = client.open_by_key(st.secrets["spreadsheet_id"])
    try:
        ws = sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound as e:
        raise KeyError(sheet_name) from e
    # value_render_option=unformatted: trae el número crudo (float/int) tal
    # como lo guarda Sheets internamente, sin pasar por el string mostrado en
    # pantalla. Es necesario porque la planilla usa formato local (coma como
    # separador decimal) y el parser por default de gspread asume formato
    # inglés (coma = separador de miles): con FORMATTED_VALUE terminaba
    # borrando la coma decimal y multiplicando importes por 10/100.
    records = ws.get_all_records(value_render_option=ValueRenderOption.unformatted)
    return pd.DataFrame.from_records(records)


def _parse_fecha_google(valor) -> pd.Timestamp:
    """Soporta fechas que Google Sheets guardó como número de serie (epoch de
    Sheets: días desde 1899-12-30), fechas como texto ("31/05/2026") y
    fechas que ya vienen como datetime (Excel local)."""
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(valor, unit="D")
    if isinstance(valor, (dt.date, dt.datetime, pd.Timestamp)):
        return pd.Timestamp(valor)
    return pd.to_datetime(valor, dayfirst=True, errors="coerce")


def rubro_de(tipo: str, categoria: str) -> str:
    """Clasifica cada movimiento según Tipo + Categoría. La categoría sola no
    alcanza: "Préstamo Galicia" existe como Cobro (desembolso) y como Gasto
    (amortización)."""
    if tipo == "Cobro":
        if categoria in CATEGORIAS_INGRESO_GESTION:
            return "ingresos"
        if categoria in CATEGORIAS_FINANCIACION:
            return "financiacion_entrada"
        return "sin_clasificar"
    if tipo == "Gasto":
        return RUBRO_GASTO.get(categoria, "sin_clasificar")
    return "sin_clasificar"


def _preparar_movimientos(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df[df["ID"].astype(str).str.strip() != ""] if "ID" in df.columns else df
    df["Fecha"] = df["Fecha"].apply(_parse_fecha_google)
    for col in ("Importe neto", "IVA"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    for col in ("Tipo", "Categoría", "Cliente/Proveedor", "Con/Sin factura", "Observación"):
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip()
    df["Importe total"] = df["Importe neto"] + df["IVA"]
    # signo: Gasto negativo, Cobro positivo; simplifica sumas de flujo de caja
    df["Importe firmado"] = df["Importe neto"].where(df["Tipo"] == "Cobro", -df["Importe neto"])
    df["Rubro"] = [rubro_de(t, c) for t, c in zip(df["Tipo"], df["Categoría"])]
    texto = df["Cliente/Proveedor"] + " " + df["Observación"]
    df["Provisoria"] = texto.str.contains(MARCAS_PROVISORIAS)
    df = df.dropna(subset=["Fecha"])
    df["Mes"] = df["Fecha"].dt.to_period("M")
    return df


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner="Leyendo la planilla…")
def load_movimientos() -> pd.DataFrame:
    if modo_demo_activo():
        return _preparar_movimientos(_demo_movimientos())
    df = _leer_hoja("Movimientos")
    if df.empty:
        return df
    return _preparar_movimientos(df)


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def load_supuestos() -> dict:
    if modo_demo_activo():
        return _demo_supuestos()
    df = _leer_hoja("Supuestos")
    return dict(zip(df["Parámetro"], df["Valor"]))


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def load_obligaciones_futuras() -> pd.DataFrame:
    if modo_demo_activo():
        return pd.DataFrame(columns=["Mes (AAAA-MM)", "Categoría", "Importe estimado", "Observación"])
    df = _leer_hoja("Obligaciones_Futuras")
    if not df.empty:
        df["Importe estimado"] = pd.to_numeric(df["Importe estimado"], errors="coerce").fillna(0)
    return df


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def load_clientes() -> dict:
    """Cantidad de clientes por línea de negocio (hoja `Clientes`). No es un
    dato transaccional, se actualiza cada tanto; se usa para Punto de
    Equilibrio y Sensibilidad."""
    if modo_demo_activo():
        return {"Liquidación de Sueldos": 48, "Tercerización de Estudios": 9,
                "Asesoramiento Laboral": 1, "Flujo Sole": 0, "Mandú": 6, "Otro ingreso": 0}
    try:
        df = _leer_hoja("Clientes")
    except KeyError:
        return {}
    if df.empty:
        return {}
    df["Cantidad de clientes"] = pd.to_numeric(df["Cantidad de clientes"], errors="coerce").fillna(0)
    return dict(zip(df["Línea de negocio"], df["Cantidad de clientes"]))


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def load_calendario() -> pd.DataFrame:
    """Compromisos de cobro/pago con fecha cierta, todavía no ocurridos (hoja
    `Calendario_Vencimientos`). A diferencia de Movimientos, esto no son
    hechos consumados: es lo que se sabe que va a pasar, cargado a mano."""
    cols = ["ID", "Fecha esperada", "Tipo", "Categoría", "Concepto", "Importe", "Estado", "Observación"]
    if modo_demo_activo():
        hoy = pd.Timestamp.today().normalize()
        return pd.DataFrame([
            {"ID": "demo-1", "Fecha esperada": hoy + pd.Timedelta(days=3), "Tipo": "Cobro",
             "Categoría": "Liquidación de Sueldos", "Concepto": "Cliente grande, factura pendiente",
             "Importe": 4500000, "Estado": "Pendiente", "Observación": ""},
            {"ID": "demo-2", "Fecha esperada": hoy + pd.Timedelta(days=5), "Tipo": "Gasto",
             "Categoría": "Nómina", "Concepto": "Sueldos del mes", "Importe": 15000000,
             "Estado": "Pendiente", "Observación": ""},
            {"ID": "demo-3", "Fecha esperada": hoy + pd.Timedelta(days=12), "Tipo": "Gasto",
             "Categoría": "Cuotas ARCA", "Concepto": "Vencimiento ARCA", "Importe": 4000000,
             "Estado": "Pendiente", "Observación": ""},
        ], columns=cols)
    try:
        df = _leer_hoja("Calendario_Vencimientos")
    except KeyError:
        return pd.DataFrame(columns=cols)
    if df.empty or "Fecha esperada" not in df.columns:
        return pd.DataFrame(columns=cols)
    df["Fecha esperada"] = df["Fecha esperada"].apply(_parse_fecha_google)
    df["Importe"] = pd.to_numeric(df["Importe"], errors="coerce").fillna(0)
    return df.dropna(subset=["Fecha esperada"])


# ──────────────────────────────────────────────────────────────────────────
# Datos de demostración (solo si no hay ninguna fuente configurada)
# ──────────────────────────────────────────────────────────────────────────


def _demo_movimientos() -> pd.DataFrame:
    import numpy as np

    rng = np.random.default_rng(42)
    hoy = pd.Timestamp.today().normalize()
    inicio = (hoy - pd.DateOffset(months=3)).replace(day=1)
    fechas = pd.date_range(inicio, hoy, freq="3D")

    filas = []
    for i, f in enumerate(fechas):
        cat_c = rng.choice(CATEGORIAS_INGRESO_GESTION, p=[0.35, 0.25, 0.05, 0.15, 0.15, 0.05])
        filas.append({
            "ID": f"{f:%Y%m%d}-{i:03d}", "Fecha": f, "Tipo": "Cobro", "Categoría": cat_c,
            "Cliente/Proveedor": f"Cliente {i % 12}",
            "Con/Sin factura": rng.choice(["Con factura", "Sin factura"], p=[0.85, 0.15]),
            "Importe neto": float(rng.uniform(300000, 3500000)), "IVA": 0.0,
            "Observación": "", "Cargado por": "demo", "Fecha de carga": f,
        })
        if i % 2 == 0:
            cat_g = rng.choice(CATEGORIAS_GASTO)
            filas.append({
                "ID": f"{f:%Y%m%d}-{i:03d}g", "Fecha": f, "Tipo": "Gasto", "Categoría": cat_g,
                "Cliente/Proveedor": f"Proveedor {i % 8}", "Con/Sin factura": "N/A",
                "Importe neto": float(rng.uniform(100000, 2500000)), "IVA": 0.0,
                "Observación": "", "Cargado por": "demo", "Fecha de carga": f,
            })
    filas.append({
        "ID": "demo-prestamo", "Fecha": inicio, "Tipo": "Cobro", "Categoría": "Préstamo Galicia",
        "Cliente/Proveedor": "Desembolso préstamo", "Con/Sin factura": "N/A",
        "Importe neto": 5000000.0, "IVA": 0.0, "Observación": "", "Cargado por": "demo",
        "Fecha de carga": inicio,
    })
    return pd.DataFrame(filas)


def _demo_supuestos() -> dict:
    return {
        "Tipo de cambio USD/ARS": "1507",
        "Inflación mensual estimada": "0.019",
        "Objetivo margen neto de gestión": "0.30",
        "Objetivo caja SAC sin financiación: fecha": "2026-12-01",
        "Objetivo caja SAC sin financiación: umbral": "0",
        "TNA préstamo Banco Galicia": "0.32",
        "Monto préstamo Banco Galicia": "5086000",
        "Plazo préstamo Banco Galicia (meses)": "6",
    }


# ──────────────────────────────────────────────────────────────────────────
# Cálculos
# ──────────────────────────────────────────────────────────────────────────


def _objetivo_margen(supuestos: dict) -> float:
    try:
        return float(supuestos.get("Objetivo margen neto de gestión", 0.30) or 0.30)
    except (TypeError, ValueError):
        return 0.30


def resumen_mes(df: pd.DataFrame, mes: pd.Period) -> dict:
    """Suma cada rubro del mes y arma el Estado de Resultados y el puente a
    caja. Es la única función que agrega importes: todo lo demás (KPIs, ER,
    evolución, punto de equilibrio) sale de acá, así no hay dos criterios."""
    m = df[df["Mes"] == mes]
    s = {r: float(m.loc[m["Rubro"] == r, "Importe neto"].sum()) for r in RUBROS}

    r = dict(s)
    r["costos_directos"] = s["nomina"] + s["costos_externos"] + s["otros_gastos"]
    r["resultado_operativo_1"] = s["ingresos"] - r["costos_directos"]
    r["resultado_operativo_2"] = r["resultado_operativo_1"] - s["estructura"]
    r["resultado_antes_ganancias"] = r["resultado_operativo_2"] - s["iibb"]
    r["impuestos"] = s["iibb"] + s["ganancias"]
    r["resultado_neto"] = r["resultado_antes_ganancias"] - s["ganancias"]
    r["egresos_gestion"] = r["costos_directos"] + s["estructura"] + r["impuestos"]
    r["margen"] = (r["resultado_neto"] / s["ingresos"]) if s["ingresos"] else 0.0
    r["financiacion_neta"] = s["financiacion_entrada"] - s["financiacion_salida"]
    r["flujo_caja"] = r["resultado_neto"] - s["arca"] - s["retiros"] + r["financiacion_neta"]
    ing = m[m["Rubro"] == "ingresos"]
    r["ingresos_sin_factura"] = float(ing.loc[ing["Con/Sin factura"] == "Sin factura", "Importe neto"].sum())
    r["movimientos"] = len(m)
    r["provisorias"] = m[m["Provisoria"]]
    return r


def kpis_del_mes(df: pd.DataFrame, mes: pd.Period) -> dict:
    """KPIs del mes, consistentes con el Estado de Resultados. El flujo de
    caja se informa aparte (incluye ARCA, retiros y financiación)."""
    r = resumen_mes(df, mes)
    return {
        "ingresos": r["ingresos"], "egresos": r["egresos_gestion"],
        "resultado": r["resultado_neto"], "margen": r["margen"],
        "flujo_caja": r["flujo_caja"], "movimientos": r["movimientos"],
    }


def calcular_er(df: pd.DataFrame, mes: pd.Period) -> dict:
    """Estado de Resultados del mes con el criterio de la hoja `ER` del Excel
    original: Cuotas ARCA, Retiros de socios y el préstamo (desembolso y
    amortización) quedan fuera del resultado de gestión."""
    r = resumen_mes(df, mes)
    r["no_operativo"] = r["arca"] + r["retiros"]
    return r


def participacion_por_linea(df: pd.DataFrame, mes: pd.Period) -> dict:
    """% que representa cada línea sobre el ingreso de gestión del mes (sin
    financiación). Es la clave de reparto del Excel original para asignar
    costos, estructura e impuestos a cada línea. Suma 100%."""
    m = df[(df["Mes"] == mes) & (df["Rubro"] == "ingresos")]
    total = m["Importe neto"].sum()
    if not total:
        return {linea: 0.0 for linea in LINEAS_DE_NEGOCIO}
    return {
        linea: m.loc[m["Categoría"] == linea, "Importe neto"].sum() / total
        for linea in LINEAS_DE_NEGOCIO
    }


def er_por_linea(df: pd.DataFrame, mes: pd.Period, supuestos: dict) -> pd.DataFrame:
    """Rentabilidad por línea de negocio, prorrateando costos directos,
    gastos de estructura e impuestos según la participación de cada línea en
    los ingresos del mes (criterio de la hoja `ER x Línea`). Nómina no está
    separada en directa/indirecta, así que se prorratea entera: es una
    simplificación consciente frente al original."""
    er = calcular_er(df, mes)
    participacion = participacion_por_linea(df, mes)
    margen_objetivo = _objetivo_margen(supuestos)

    filas = []
    for linea in LINEAS_DE_NEGOCIO:
        p = participacion[linea]
        ingresos_linea = er["ingresos"] * p
        if not ingresos_linea:
            continue
        costos_directos_linea = er["costos_directos"] * p
        gastos_estructura_linea = er["estructura"] * p
        impuestos_linea = er["impuestos"] * p
        resultado_linea = ingresos_linea - costos_directos_linea - gastos_estructura_linea - impuestos_linea
        margen_linea = resultado_linea / ingresos_linea
        filas.append({
            "Línea": linea, "Participación": p, "Ingresos": ingresos_linea,
            "Costos directos": -costos_directos_linea, "Gastos de estructura": -gastos_estructura_linea,
            "Impuestos": -impuestos_linea, "Resultado neto": resultado_linea,
            "Margen neto %": margen_linea, "Objetivo": margen_objetivo,
            "Cumple": margen_linea >= margen_objetivo,
        })
    return pd.DataFrame(filas)


def punto_equilibrio(df: pd.DataFrame, mes: pd.Period, clientes: dict) -> dict:
    """Estructura a cubrir: costos directos + estructura + Cuotas ARCA +
    amortización del préstamo (estas dos no son costo contable pero sí
    salidas de caja obligatorias), más impuestos. Y cuántos clientes por
    línea harían falta para cubrirla con el precio promedio actual."""
    er = calcular_er(df, mes)
    estructura_operativa = er["costos_directos"] + er["estructura"] + er["arca"] + er["financiacion_salida"]
    estructura_total = estructura_operativa + er["impuestos"]
    margen_seguridad = er["ingresos"] - estructura_total
    margen_seguridad_pct = (margen_seguridad / er["ingresos"]) if er["ingresos"] else 0.0

    participacion = participacion_por_linea(df, mes)
    filas_lineas = []
    if clientes:
        for linea in LINEAS_DE_NEGOCIO:
            p = participacion[linea]
            clientes_actuales = clientes.get(linea, 0)
            if not clientes_actuales:
                continue
            ingresos_linea = er["ingresos"] * p
            precio_promedio = ingresos_linea / clientes_actuales
            estructura_asignada = estructura_total * p
            clientes_necesarios = (estructura_asignada / precio_promedio) if precio_promedio else None
            filas_lineas.append({
                "Línea": linea, "Clientes actuales": clientes_actuales,
                "Precio promedio/cliente": precio_promedio,
                "Clientes necesarios (si cubriera toda la estructura sola)": clientes_necesarios,
            })

    return {
        "estructura_operativa": estructura_operativa,
        "estructura_total": estructura_total,
        "margen_seguridad": margen_seguridad,
        "margen_seguridad_pct": margen_seguridad_pct,
        "por_linea": pd.DataFrame(filas_lineas) if filas_lineas else pd.DataFrame(),
    }


def sensibilidad_resultado(df: pd.DataFrame, mes: pd.Period, var_clientes: float, var_honorarios: float) -> float:
    """Resultado de gestión mensual (antes de Ganancias) si la cantidad de
    clientes variara `var_clientes` y el honorario promedio `var_honorarios`.
    Replica la fórmula de la hoja `Sensibilidad` del Excel original:

        Resultado = Ingresos_de_gestión × (1+honorarios) × (1+clientes)
                    × (1 − Ratio_facturado × Tasa_IIBB_efectiva)
                    − Costos_fijos_operativos

    Ingresos de gestión = sin financiación. Ratio facturado = ingresos sin
    Mandú sobre el total. Costos fijos = costos directos + estructura, que no
    escalan con el escenario."""
    er = calcular_er(df, mes)
    ingresos_base = er["ingresos"]
    costos_fijos_op = er["costos_directos"] + er["estructura"]

    m = df[(df["Mes"] == mes) & (df["Rubro"] == "ingresos")]
    facturacion_operativa = m.loc[m["Categoría"] != "Mandú", "Importe neto"].sum()
    ratio_facturado = (facturacion_operativa / ingresos_base) if ingresos_base else 0.0
    tasa_iibb_efectiva = (er["iibb"] / facturacion_operativa) if facturacion_operativa else 0.0

    ingresos_aj = ingresos_base * (1 + var_honorarios) * (1 + var_clientes) * (1 - ratio_facturado * tasa_iibb_efectiva)
    return ingresos_aj - costos_fijos_op


def matriz_sensibilidad(df: pd.DataFrame, mes: pd.Period,
                         pasos=(-0.2, -0.1, 0.0, 0.1, 0.2)) -> pd.DataFrame:
    """Filas = variación en cantidad de clientes (Q), columnas = variación en
    honorario promedio (P), mismo layout que el Excel original."""
    filas = []
    for vq in pasos:
        fila = {"Variación clientes": vq}
        for vp in pasos:
            fila[vp] = sensibilidad_resultado(df, mes, vq, vp)
        filas.append(fila)
    return pd.DataFrame(filas).set_index("Variación clientes")


def desglose_por_categoria(df: pd.DataFrame, mes: pd.Period, rubros: list) -> pd.DataFrame:
    m = df[(df["Mes"] == mes) & (df["Rubro"].isin(rubros))]
    return (
        m.groupby("Categoría")["Importe neto"].sum().sort_values(ascending=False).reset_index()
    )


def meses_con_datos(df: pd.DataFrame) -> pd.PeriodIndex:
    if df.empty:
        return pd.PeriodIndex([], freq="M")
    return pd.period_range(df["Mes"].min(), df["Mes"].max(), freq="M")


def evolucion_mensual(df: pd.DataFrame) -> pd.DataFrame:
    """Una fila por mes con las líneas del ER y el puente a caja."""
    filas = []
    for m in meses_con_datos(df):
        r = resumen_mes(df, m)
        filas.append({
            "Mes": m.strftime("%m/%Y"), "Periodo": m,
            "Ingresos de gestión": r["ingresos"],
            "Nómina": r["nomina"], "Costos externos": r["costos_externos"],
            "Otros gastos": r["otros_gastos"], "Gastos de estructura": r["estructura"],
            "IIBB": r["iibb"], "Ganancias": r["ganancias"],
            "Egresos de gestión": r["egresos_gestion"],
            "Resultado neto de gestión": r["resultado_neto"], "Margen neto": r["margen"],
            "Cuotas ARCA": r["arca"], "Retiros de socios": r["retiros"],
            "Financiación neta": r["financiacion_neta"], "Flujo de caja": r["flujo_caja"],
            "Partidas provisorias": len(r["provisorias"]),
        })
    return pd.DataFrame(filas)


def evolucion_por_categoria(df: pd.DataFrame) -> pd.DataFrame:
    """Matriz categoría × mes (importe neto), separando ingresos de gestión,
    financiación y gastos."""
    if df.empty:
        return pd.DataFrame()
    d = df.copy()
    etiquetas = {"ingresos": "Ingreso", "financiacion_entrada": "Financiación (entrada)",
                 "financiacion_salida": "Financiación (salida)"}
    d["Bloque"] = d["Rubro"].map(etiquetas).fillna("Gasto")
    t = d.pivot_table(index=["Bloque", "Categoría"], columns="Mes", values="Importe neto",
                      aggfunc="sum", fill_value=0)
    t = t.reindex(columns=meses_con_datos(df), fill_value=0)
    t.columns = [c.strftime("%m/%Y") for c in t.columns]
    orden = {"Ingreso": 0, "Gasto": 1, "Financiación (entrada)": 2, "Financiación (salida)": 3}
    t = t.reset_index()
    t["_o"] = t["Bloque"].map(orden)
    return t.sort_values(["_o", "Categoría"]).drop(columns="_o").set_index(["Bloque", "Categoría"])


def controles_de_carga(df: pd.DataFrame) -> pd.DataFrame:
    """Filas que el dashboard no puede clasificar: Tipo distinto de Cobro o
    Gasto, categoría inexistente o combinación inválida (por ejemplo un Cobro
    con categoría Nómina). Esas filas no entran en ningún cálculo."""
    return df[df["Rubro"] == "sin_clasificar"]


def flujo_caja_real(df: pd.DataFrame) -> pd.DataFrame:
    """Caja acumulada mes a mes según lo cargado: resultado de gestión menos
    ARCA y retiros, más la financiación neta. No es un saldo bancario: incluye
    cobros en efectivo y lo que cobra Sole, y los impuestos van por lo
    devengado."""
    filas = []
    caja = 0.0
    for m in meses_con_datos(df):
        caja += resumen_mes(df, m)["flujo_caja"]
        filas.append({"mes": m.to_timestamp(), "caja": caja})
    return pd.DataFrame(filas, columns=["mes", "caja"])


def calendario_liquidez(df_calendario: pd.DataFrame, caja_actual: float, dias_adelante: int = 90) -> pd.DataFrame:
    """Calce de liquidez de corto plazo: ordena los compromisos PENDIENTES del
    calendario dentro de la ventana de `dias_adelante` días desde hoy y va
    acumulando el saldo desde `caja_actual`. Es un calce de fechas, no una
    proyección."""
    hoy = pd.Timestamp.today().normalize()
    limite = hoy + pd.Timedelta(days=dias_adelante)
    pendientes = df_calendario[
        (df_calendario.get("Estado", "Pendiente") == "Pendiente")
        & (df_calendario["Fecha esperada"] >= hoy) & (df_calendario["Fecha esperada"] <= limite)
    ].sort_values("Fecha esperada").copy()

    pendientes["Importe firmado"] = pendientes["Importe"].where(pendientes["Tipo"] == "Cobro", -pendientes["Importe"])
    pendientes["Saldo proyectado"] = caja_actual + pendientes["Importe firmado"].cumsum()
    return pendientes


def vencidos_pendientes(df_calendario: pd.DataFrame) -> pd.DataFrame:
    """Compromisos cuya fecha ya pasó pero siguen en 'Pendiente': señal de
    que probablemente ya sucedieron y falta marcarlos como 'Cumplido'. Es un
    recordatorio, no una corrección automática."""
    if df_calendario.empty:
        return df_calendario
    hoy = pd.Timestamp.today().normalize()
    return df_calendario[
        (df_calendario.get("Estado", "Pendiente") == "Pendiente")
        & (df_calendario["Fecha esperada"] < hoy)
    ].sort_values("Fecha esperada")


def evaluar_semaforo(df: pd.DataFrame, supuestos: dict, mes: pd.Period) -> list:
    """Objetivos evaluados con datos del mes que se está mirando."""
    margen_obj = _objetivo_margen(supuestos)
    r = resumen_mes(df, mes)
    return [{
        "nombre": f"Margen neto de gestión del mes ≥ {margen_obj:.0%}",
        "valor": porcentaje(r["margen"]), "cumple": r["margen"] >= margen_obj,
    }]


def _delta(actual: float, anterior, como_pct: bool = False):
    if anterior is None:
        return None
    if como_pct:
        return f"{(actual - anterior) * 100:+.1f} p.p.".replace(".", ",", 1)
    if not anterior:
        return None
    return f"{porcentaje((actual - anterior) / abs(anterior), signo=True)} vs mes anterior"


# ──────────────────────────────────────────────────────────────────────────
# Interfaz
# ──────────────────────────────────────────────────────────────────────────

st.title("📊 Estudio Haberes: dashboard financiero en tiempo real")

fuente = fuente_de_datos()
if fuente == "demo":
    st.warning(
        "**MODO DEMO**: no se encontraron credenciales de Google Sheets en `st.secrets` "
        "ni la variable HABERES_XLSX. Mostrando datos sintéticos para probar la interfaz. "
        "Configurá `gcp_service_account` y `spreadsheet_id` según SETUP_GOOGLE_CLOUD.md "
        "para conectar la planilla real.",
        icon="⚠️",
    )
elif fuente == "excel":
    st.info(f"Leyendo el Excel local {os.environ['HABERES_XLSX']} (variable HABERES_XLSX).", icon="📄")

col_refresh, _ = st.columns([1, 5])
if col_refresh.button("🔄 Actualizar ahora"):
    st.cache_data.clear()
    st.rerun()

df_mov = load_movimientos()
supuestos = load_supuestos()

if df_mov.empty:
    st.info("Todavía no hay movimientos cargados en la planilla.")
    st.stop()

# ── Filtros (sidebar) ──────────────────────────────────────────────────
st.sidebar.header("Filtros")
meses_disponibles = sorted(df_mov["Mes"].unique(), reverse=True)
mes_kpi = st.sidebar.selectbox("Mes de referencia", meses_disponibles, index=0,
                               format_func=lambda p: p.strftime("%m/%Y"))
st.sidebar.caption("El mes de referencia controla indicadores, estado de resultados y análisis. "
                   "Los filtros de abajo solo afectan la tabla de movimientos.")
fecha_min, fecha_max = df_mov["Fecha"].min().date(), df_mov["Fecha"].max().date()
rango = st.sidebar.date_input("Rango de fechas", (fecha_min, fecha_max),
                               min_value=fecha_min, max_value=fecha_max)
categorias_todas = sorted(df_mov["Categoría"].dropna().unique())
cat_sel = st.sidebar.multiselect("Categorías", categorias_todas, default=categorias_todas)
tipo_sel = st.sidebar.multiselect("Tipo", ["Cobro", "Gasto"], default=["Cobro", "Gasto"])

if isinstance(rango, tuple) and len(rango) == 2:
    desde, hasta = rango
    df_filtrado = df_mov[
        (df_mov["Fecha"].dt.date >= desde) & (df_mov["Fecha"].dt.date <= hasta)
        & (df_mov["Categoría"].isin(cat_sel)) & (df_mov["Tipo"].isin(tipo_sel))
    ]
else:
    df_filtrado = df_mov[(df_mov["Categoría"].isin(cat_sel)) & (df_mov["Tipo"].isin(tipo_sel))]

# ── Control de calidad de la carga ─────────────────────────────────────
sin_clasificar = controles_de_carga(df_mov)
if not sin_clasificar.empty:
    st.error(
        f"Hay {len(sin_clasificar)} fila(s) en Movimientos con Tipo o Categoría que el dashboard "
        "no reconoce. No entran en ningún cálculo hasta corregirlas.", icon="🚫"
    )
    st.dataframe(sin_clasificar[["ID", "Fecha", "Tipo", "Categoría", "Cliente/Proveedor", "Importe neto"]],
                 width="stretch", hide_index=True)

r_mes = resumen_mes(df_mov, mes_kpi)
mes_anterior = mes_kpi - 1
r_ant = resumen_mes(df_mov, mes_anterior) if mes_anterior in set(df_mov["Mes"]) else None

if not r_mes["provisorias"].empty:
    prov = r_mes["provisorias"]
    st.markdown(
        f'<span class="semaforo-badge semaforo-aviso">⚠ {mes_kpi.strftime("%m/%Y")} tiene '
        f'{len(prov)} partida(s) provisoria(s)</span>'
        f'<span style="color:{GRIS_TEXTO}">Los resultados de este mes pueden cambiar cuando se completen.</span>',
        unsafe_allow_html=True,
    )
    with st.expander("Ver partidas provisorias"):
        st.dataframe(
            prov[["ID", "Categoría", "Cliente/Proveedor", "Importe neto", "Observación"]]
            .style.format({"Importe neto": pesos}),
            width="stretch", hide_index=True,
        )

# ── KPIs del mes ───────────────────────────────────────────────────────
st.header(f"Indicadores de {mes_kpi.strftime('%m/%Y')}")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Ingresos de gestión", pesos(r_mes["ingresos"]),
          _delta(r_mes["ingresos"], r_ant and r_ant["ingresos"]))
c2.metric("Egresos de gestión", pesos(r_mes["egresos_gestion"]),
          _delta(r_mes["egresos_gestion"], r_ant and r_ant["egresos_gestion"]), delta_color="inverse")
c3.metric("Resultado neto de gestión", pesos(r_mes["resultado_neto"]),
          _delta(r_mes["resultado_neto"], r_ant and r_ant["resultado_neto"]))
c4.metric("Margen neto", porcentaje(r_mes["margen"]),
          _delta(r_mes["margen"], r_ant["margen"] if r_ant else None, como_pct=True))
c5, c6, c7, c8 = st.columns(4)
c5.metric("Flujo de caja del mes", pesos(r_mes["flujo_caja"]))
c6.metric("Retiros de socios", pesos(r_mes["retiros"]))
c7.metric("Cuotas ARCA", pesos(r_mes["arca"]))
c8.metric("Financiación neta (préstamo)", pesos(r_mes["financiacion_neta"]))
if r_mes["ingresos"]:
    st.caption(f"Ingresos sin factura: {pesos(r_mes['ingresos_sin_factura'])} "
               f"({porcentaje(r_mes['ingresos_sin_factura'] / r_mes['ingresos'])} del ingreso de gestión).")

# ── Estado de Resultados ───────────────────────────────────────────────
st.subheader(f"Estado de Resultados de {mes_kpi.strftime('%m/%Y')}")
er = calcular_er(df_mov, mes_kpi)
filas_er = [
    ("Ingresos de gestión", er["ingresos"]),
    ("(−) Nómina", -er["nomina"]),
    ("(−) Costos externos", -er["costos_externos"]),
    ("(−) Otros gastos (incluye intereses del préstamo)", -er["otros_gastos"]),
    ("= Resultado operativo (antes de estructura)", er["resultado_operativo_1"]),
    ("(−) Gastos de estructura", -er["estructura"]),
    ("= Resultado operativo (antes de impuestos)", er["resultado_operativo_2"]),
    ("(−) Ingresos Brutos", -er["iibb"]),
    ("= Resultado antes de Ganancias", er["resultado_antes_ganancias"]),
    ("(−) Impuesto a las Ganancias", -er["ganancias"]),
    ("= Resultado neto de gestión", er["resultado_neto"]),
]
df_er = pd.DataFrame(filas_er, columns=["Línea", "Importe"])
df_er["% s/ingresos"] = df_er["Importe"] / er["ingresos"] if er["ingresos"] else 0.0

col_tabla, col_donut = st.columns([3, 2])
with col_tabla:
    st.dataframe(
        df_er.style.format({"Importe": pesos, "% s/ingresos": porcentaje}),
        width="stretch", hide_index=True,
    )
with col_donut:
    composicion = pd.DataFrame({
        "Componente": ["Nómina", "Costos externos y otros", "Gastos de estructura", "Impuestos", "Resultado neto"],
        "Importe": [er["nomina"], er["costos_externos"] + er["otros_gastos"], er["estructura"],
                    er["impuestos"], max(er["resultado_neto"], 0)],
    })
    composicion = composicion[composicion["Importe"] > 0]
    fig_donut = go.Figure(data=[go.Pie(
        labels=composicion["Componente"], values=composicion["Importe"], hole=0.55,
        marker_colors=[ROJO_SELLO, "#D9737F", DORADO_SELLO, "#8A8168", VERDE_LEDGER],
        sort=False,
    )])
    plotly_layout_base(fig_donut, showlegend=True, legend=dict(orientation="h", y=-0.1))
    fig_donut.update_traces(textinfo="percent")
    st.plotly_chart(fig_donut, width="stretch")

st.markdown("**Del resultado de gestión al flujo de caja del mes**")
puente = pd.DataFrame([
    ("Resultado neto de gestión", er["resultado_neto"]),
    ("(−) Cuotas planes ARCA", -er["arca"]),
    ("(−) Retiros de socios", -er["retiros"]),
    ("(+) Préstamo recibido", er["financiacion_entrada"]),
    ("(−) Amortización de capital del préstamo", -er["financiacion_salida"]),
    ("= Flujo de caja del mes", er["flujo_caja"]),
], columns=["Concepto", "Importe"])
st.dataframe(puente.style.format({"Importe": pesos}), width="stretch", hide_index=True)
st.caption(
    "ARCA, retiros y préstamo no son costo de gestión (mismo criterio que la hoja ER del Excel "
    "original) pero sí mueven la caja. El préstamo recibido no es ingreso: es financiación."
)

# ── Evolución mensual ───────────────────────────────────────────────────
st.subheader("Evolución mensual")
evo = evolucion_mensual(df_mov)
fig_evo = go.Figure()
fig_evo.add_trace(go.Bar(x=evo["Mes"], y=evo["Ingresos de gestión"], name="Ingresos de gestión",
                         marker_color=VERDE_LEDGER))
fig_evo.add_trace(go.Bar(x=evo["Mes"], y=evo["Egresos de gestión"], name="Egresos de gestión",
                         marker_color=ROJO_SELLO))
fig_evo.add_trace(go.Scatter(x=evo["Mes"], y=evo["Resultado neto de gestión"], name="Resultado neto",
                             mode="lines+markers", line=dict(color=INK, width=2)))
fig_evo.add_trace(go.Scatter(x=evo["Mes"], y=evo["Flujo de caja"], name="Flujo de caja",
                             mode="lines+markers", line=dict(color=DORADO_SELLO, width=2, dash="dot")))
plotly_layout_base(fig_evo, barmode="group", legend=dict(orientation="h", y=-0.15), yaxis_title="$")
st.plotly_chart(fig_evo, width="stretch")

cols_evo = ["Mes", "Ingresos de gestión", "Nómina", "Costos externos", "Otros gastos",
            "Gastos de estructura", "IIBB", "Ganancias", "Resultado neto de gestión", "Margen neto",
            "Cuotas ARCA", "Retiros de socios", "Financiación neta", "Flujo de caja", "Partidas provisorias"]
fmt_evo = {c: pesos for c in cols_evo if c not in ("Mes", "Margen neto", "Partidas provisorias")}
fmt_evo["Margen neto"] = porcentaje
st.dataframe(evo[cols_evo].style.format(fmt_evo), width="stretch", hide_index=True)
if evo["Partidas provisorias"].any():
    st.caption("Los meses con partidas provisorias (marcadas PARCIAL, PENDIENTE, ESTIMADO o A CONFIRMAR en "
               "la carga) todavía no son comparables del todo con los meses cerrados.")

with st.expander("Detalle por categoría y variación contra el mes anterior"):
    por_cat = evolucion_por_categoria(df_mov)
    if not por_cat.empty:
        st.dataframe(por_cat.style.format(pesos), width="stretch")
        if por_cat.shape[1] >= 2:
            ult, ant = por_cat.columns[-1], por_cat.columns[-2]
            var = pd.DataFrame({
                ant: por_cat[ant], ult: por_cat[ult],
                "Variación $": por_cat[ult] - por_cat[ant],
                "Variación %": (por_cat[ult] - por_cat[ant]) / por_cat[ant].where(por_cat[ant] != 0),
            })
            st.markdown(f"**{ult} contra {ant}**")
            st.dataframe(var.style.format({ant: pesos, ult: pesos, "Variación $": pesos,
                                           "Variación %": lambda v: porcentaje(v, signo=True)}),
                         width="stretch")

# ── ER x Línea de negocio ───────────────────────────────────────────────
st.subheader(f"Rentabilidad por línea de negocio de {mes_kpi.strftime('%m/%Y')}")
df_er_linea = er_por_linea(df_mov, mes_kpi, supuestos)
if df_er_linea.empty:
    st.info("No hay ingresos de gestión en el mes seleccionado.")
else:
    df_er_linea_mostrar = df_er_linea.copy()
    df_er_linea_mostrar["Cumple"] = df_er_linea_mostrar["Cumple"].map({True: "✓", False: "✕"})
    st.dataframe(
        df_er_linea_mostrar.style.format({
            "Participación": porcentaje, "Ingresos": pesos, "Costos directos": pesos,
            "Gastos de estructura": pesos, "Impuestos": pesos, "Resultado neto": pesos,
            "Margen neto %": porcentaje, "Objetivo": "{:.0%}",
        }),
        width="stretch", hide_index=True,
    )
    fig_margen_linea = go.Figure()
    colores_margen = [VERDE_LEDGER if c else ROJO_SELLO for c in df_er_linea["Cumple"]]
    fig_margen_linea.add_trace(go.Bar(
        y=df_er_linea["Línea"], x=df_er_linea["Margen neto %"], orientation="h",
        marker_color=colores_margen, text=[porcentaje(v) for v in df_er_linea["Margen neto %"]],
        textposition="outside", name="Margen real",
    ))
    objetivo_linea = df_er_linea["Objetivo"].iloc[0]
    fig_margen_linea.add_vline(x=objetivo_linea, line_dash="dash", line_color=DORADO_SELLO,
                                annotation_text=f"Objetivo {objetivo_linea:.0%}")
    plotly_layout_base(fig_margen_linea, xaxis_tickformat=".0%", showlegend=False,
                        xaxis_title="Margen neto %", yaxis_title=None)
    st.plotly_chart(fig_margen_linea, width="stretch")
    st.caption(
        "Los costos se reparten entre líneas según su participación en el ingreso de gestión del mes "
        "(mismo criterio que el Excel original). Como 'Nómina' es una sola categoría (no separada en "
        "directa e indirecta), todas las líneas quedan con el mismo margen: el reparto es una "
        "aproximación hasta que se separe esa categoría."
    )

# ── Punto de equilibrio ─────────────────────────────────────────────────
st.subheader(f"Punto de equilibrio y margen de seguridad de {mes_kpi.strftime('%m/%Y')}")
clientes = load_clientes()
peq = punto_equilibrio(df_mov, mes_kpi, clientes)
c1, c2, c3 = st.columns(3)
c1.metric("Estructura total a cubrir", pesos(peq["estructura_total"]))
c2.metric("Margen de seguridad", pesos(peq["margen_seguridad"]))
c3.metric("Margen de seguridad %", porcentaje(peq["margen_seguridad_pct"]))
st.caption("Estructura = costos directos + estructura + impuestos + cuotas ARCA + amortización del préstamo.")
if not peq["por_linea"].empty:
    st.dataframe(
        peq["por_linea"].style.format({
            "Precio promedio/cliente": pesos,
            "Clientes necesarios (si cubriera toda la estructura sola)": "{:,.1f}",
        }),
        width="stretch", hide_index=True,
    )
else:
    st.info(
        "No hay datos en la hoja 'Clientes' todavía: corré extend_sheet_clientes.py "
        "para agregarla y ver el punto de equilibrio en cantidad de clientes por línea."
    )

# ── Semáforo de objetivos ──────────────────────────────────────────────
st.subheader("Semáforo de objetivos")
for o in evaluar_semaforo(df_mov, supuestos, mes_kpi):
    clase = "semaforo-ok" if o["cumple"] else "semaforo-no"
    icono = "✓" if o["cumple"] else "✕"
    st.markdown(
        f'<span class="semaforo-badge {clase}">{icono} {o["nombre"]}</span> '
        f'<span style="font-family:IBM Plex Mono, monospace;">{o["valor"]}</span>',
        unsafe_allow_html=True,
    )

# ── Flujo de caja acumulado (real) ──────────────────────────────────────
st.subheader("Flujo de caja acumulado según la carga")
serie_caja = flujo_caja_real(df_mov)
fig_caja = go.Figure()
fig_caja.add_trace(go.Scatter(x=serie_caja["mes"], y=serie_caja["caja"], mode="lines+markers",
                               name="Real", line=dict(color=INK, width=2),
                               fill="tozeroy", fillcolor=f"{INK}11"))
plotly_layout_base(fig_caja, showlegend=False, yaxis_title="Caja acumulada ($)")
st.plotly_chart(fig_caja, width="stretch")
st.caption(
    "Acumula, mes a mes, el resultado de gestión menos ARCA y retiros, más la financiación neta. "
    "No es el saldo bancario: incluye cobros en efectivo y lo que cobra Sole, y los impuestos van "
    "por lo devengado. La proyección a futuro queda a cargo del informe mensual de Épsilon."
)

# ── Calendario de liquidez (corto plazo) ────────────────────────────────
st.subheader("Calendario de liquidez: próximos 90 días")
df_calendario = load_calendario()

vencidos = vencidos_pendientes(df_calendario)
if not vencidos.empty:
    st.markdown(
        f'<span class="semaforo-badge semaforo-no">⚠ {len(vencidos)} vencimiento(s) pasado(s) '
        f'de fecha y todavía en "Pendiente"</span>',
        unsafe_allow_html=True,
    )
    st.dataframe(
        vencidos[["Fecha esperada", "Tipo", "Categoría", "Concepto", "Importe"]].style.format({
            "Fecha esperada": lambda d: d.strftime("%d/%m/%Y"), "Importe": pesos,
        }),
        width="stretch", hide_index=True,
    )
    st.caption(
        "Revisá si esto ya sucedió: si es así, marcalo como 'Cumplido' en Calendario_Vencimientos "
        "y cargalo en Movimientos si corresponde. Si el vencimiento se corrió, actualizá la fecha."
    )

caja_actual = serie_caja["caja"].iloc[-1] if not serie_caja.empty else 0.0
calce = calendario_liquidez(df_calendario, caja_actual, dias_adelante=90)

st.metric("Caja actual (según movimientos cargados)", pesos(caja_actual))

if calce.empty:
    st.info(
        "No hay compromisos pendientes cargados en 'Calendario_Vencimientos' para los próximos 90 "
        "días. Cargá ahí los vencimientos y cobros esperados con fecha."
    )
else:
    quiebre = calce[calce["Saldo proyectado"] < 0]
    if not quiebre.empty:
        primer_quiebre = quiebre.iloc[0]
        st.markdown(
            f'<span class="semaforo-badge semaforo-no">⚠ Riesgo de faltante de caja el '
            f'{primer_quiebre["Fecha esperada"].strftime("%d/%m/%Y")}</span> '
            f'<span style="font-family:IBM Plex Mono, monospace;">'
            f'saldo proyectado {pesos(primer_quiebre["Saldo proyectado"])}</span>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<span class="semaforo-badge semaforo-ok">✓ Sin faltantes de caja proyectados '
            'en los próximos 90 días</span>', unsafe_allow_html=True,
        )

    fig_calce = go.Figure()
    colores_calce = [ROJO_SELLO if v < 0 else INK for v in calce["Saldo proyectado"]]
    fig_calce.add_trace(go.Scatter(
        x=calce["Fecha esperada"], y=calce["Saldo proyectado"], mode="lines+markers",
        line=dict(color=INK, width=2), marker=dict(color=colores_calce, size=8),
    ))
    fig_calce.add_hline(y=0, line_dash="dash", line_color=ROJO_SELLO)
    plotly_layout_base(fig_calce, showlegend=False, yaxis_title="Saldo proyectado ($)")
    st.plotly_chart(fig_calce, width="stretch")

    calce_mostrar = calce[["Fecha esperada", "Tipo", "Categoría", "Concepto", "Importe firmado", "Saldo proyectado"]].copy()

    def _resaltar_negativos(fila):
        color = f"background-color: {ROJO_SELLO}22" if fila["Saldo proyectado"] < 0 else ""
        return [color] * len(fila)

    st.dataframe(
        calce_mostrar.style.apply(_resaltar_negativos, axis=1).format({
            "Fecha esperada": lambda d: d.strftime("%d/%m/%Y"),
            "Importe firmado": pesos, "Saldo proyectado": pesos,
        }),
        width="stretch", hide_index=True,
    )
    st.caption(
        "Calce de compromisos con fecha cierta cargados en 'Calendario_Vencimientos', no un "
        "pronóstico: muestra si el orden en que entran los cobros y salen los pagos deja la caja "
        "en negativo en algún momento, aunque el resultado del mes sea positivo."
    )

# ── Análisis de sensibilidad ────────────────────────────────────────────
st.subheader(f"Análisis de sensibilidad de {mes_kpi.strftime('%m/%Y')}")
matriz = matriz_sensibilidad(df_mov, mes_kpi)
total_clientes_base = sum(clientes.values()) if clientes else None
etiquetas_y = []
for vq in matriz.index:
    if total_clientes_base:
        clientes_aj = round(total_clientes_base * (1 + vq))
        etiquetas_y.append(f"{vq:+.0%} ({clientes_aj} cli.)")
    else:
        etiquetas_y.append(f"{vq:+.0%}")
fig_sens = go.Figure(data=go.Heatmap(
    z=matriz.values,
    x=[f"{c:+.0%}" for c in matriz.columns],
    y=etiquetas_y,
    text=[[pesos(v) for v in fila] for fila in matriz.values],
    texttemplate="%{text}",
    colorscale="RdYlGn",
    zmid=0,
))
fig_sens.update_layout(
    xaxis_title="Variación en honorario promedio (P)",
    yaxis_title="Variación en cantidad de clientes (Q)",
)
plotly_layout_base(fig_sens)
st.plotly_chart(fig_sens, width="stretch")
st.caption(
    "Resultado de gestión mensual (antes de Ganancias) según cómo varíen la cantidad de clientes (Q) "
    "y el honorario promedio (P), con los costos operativos fijos. Misma lógica que la tabla de "
    "sensibilidad del Excel original; el préstamo no entra como ingreso."
)

# ── Desglose por categoría ─────────────────────────────────────────────
col_a, col_b = st.columns(2)
with col_a:
    st.subheader("Ingresos de gestión por línea")
    df_ing = desglose_por_categoria(df_mov, mes_kpi, ["ingresos"])
    fig_ing = go.Figure(go.Bar(x=df_ing["Importe neto"], y=df_ing["Categoría"], orientation="h",
                                marker_color=VERDE_LEDGER))
    plotly_layout_base(fig_ing, yaxis=dict(autorange="reversed"), xaxis_title="$")
    st.plotly_chart(fig_ing, width="stretch")
with col_b:
    st.subheader("Egresos por categoría")
    rubros_gasto = [r for r in RUBROS if r not in ("ingresos", "financiacion_entrada")]
    df_gas = desglose_por_categoria(df_mov, mes_kpi, rubros_gasto)
    fig_gas = go.Figure(go.Bar(x=df_gas["Importe neto"], y=df_gas["Categoría"], orientation="h",
                                marker_color=ROJO_SELLO))
    plotly_layout_base(fig_gas, yaxis=dict(autorange="reversed"), xaxis_title="$")
    st.plotly_chart(fig_gas, width="stretch")
    st.caption("Incluye partidas fuera del resultado de gestión (ARCA, retiros, amortización del préstamo).")

# ── Tabla de movimientos filtrada ──────────────────────────────────────
st.subheader("Movimientos (según filtros)")
columnas_tabla = [c for c in df_filtrado.columns if c not in ("Importe firmado", "Rubro", "Provisoria", "Mes")]
st.dataframe(df_filtrado.sort_values("Fecha", ascending=False)[columnas_tabla], width="stretch")
st.download_button(
    "Descargar CSV filtrado", df_filtrado[columnas_tabla].to_csv(index=False).encode("utf-8"),
    file_name="movimientos_filtrados.csv", mime="text/csv",
)
