"""
Dashboard financiero en tiempo real — Estudio Haberes.

Lee el libro diario (hoja "Movimientos") y los parámetros de proyección
(hojas "Supuestos" y "Obligaciones_Futuras") de una planilla de Google Sheets,
vía la API de Google con una cuenta de servicio.

Carga de datos: SOLO directa en Google Sheets (no hay formulario de carga
dentro de esta app, por decisión explícita del cliente/consultor).

Correr con:
    streamlit run app.py

Requiere los secrets `gcp_service_account` y `spreadsheet_id`
(ver SETUP_GOOGLE_CLOUD.md). Si no están configurados, la app arranca en
MODO DEMO con datos sintéticos, para poder probar la interfaz sin credenciales.
"""

import datetime as dt

import gspread
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.utils import ValueRenderOption

# ──────────────────────────────────────────────────────────────────────────
# Configuración y categorías (deben coincidir con setup_sheets.py)
# ──────────────────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
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

# Mapeo de categoría → línea del Estado de Resultados (mismo criterio que la
# hoja `ER` del Excel original). "Nómina" acá viene como una sola categoría
# (directa + indirecta combinada) porque así se carga en Movimientos — si más
# adelante hace falta separarlas, se puede agregar "Nómina directa" /
# "Nómina indirecta" como categorías nuevas y mapearlas a distintos buckets.
CATEGORIA_A_LINEA_ER = {
    **{c: "ingresos" for c in CATEGORIAS_COBRO},
    "Nómina": "costos_directos",
    "Costos externos": "costos_directos",
    "Otro gasto": "costos_directos",
    "Gastos de estructura": "gastos_estructura",
    "Impuestos IIBB": "impuestos",
    "Impuestos Ganancias": "impuestos",
    "Cuotas ARCA": "no_operativo",
    "Retiros de socios": "no_operativo",
    "Préstamo Galicia": "no_operativo",
}

CACHE_TTL_SEGUNDOS = 60

# Paleta y tipografía — pensada para el mundo de un estudio contable (libro
# diario, sellos, papel), no una paleta genérica de IA. Los mismos colores se
# usan tanto en el CSS como en los gráficos de Plotly más abajo.
INK = "#1F2D4A"        # tinta / texto principal
PAPER = "#F6F4EE"      # fondo, papel
CARD = "#FFFFFF"
VERDE_LEDGER = "#3B7A57"   # cumple objetivo
ROJO_SELLO = "#B23A48"     # no cumple objetivo
DORADO_SELLO = "#C08A28"   # acento / advertencia
GRIS_TEXTO = "#6B6558"     # texto secundario

st.set_page_config(page_title="Estudio Haberes — Dashboard", layout="wide")

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
/* Números en monoespaciada — como columnas de un libro contable */
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

# ──────────────────────────────────────────────────────────────────────────
# Conexión a Google Sheets
# ──────────────────────────────────────────────────────────────────────────


def modo_demo_activo() -> bool:
    return "gcp_service_account" not in st.secrets or "spreadsheet_id" not in st.secrets


@st.cache_resource(show_spinner=False)
def get_client() -> gspread.Client:
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )
    return gspread.authorize(creds)


def _worksheet_to_df(client: gspread.Client, sheet_name: str) -> pd.DataFrame:
    sh = client.open_by_key(st.secrets["spreadsheet_id"])
    ws = sh.worksheet(sheet_name)
    # value_render_option=unformatted: trae el número crudo (float/int) tal
    # como lo guarda Sheets internamente, sin pasar por el string mostrado en
    # pantalla. Es necesario porque la planilla usa formato local (coma como
    # separador decimal) y el parser por default de gspread asume formato
    # inglés (coma = separador de miles) — con FORMATTED_VALUE terminaba
    # borrando la coma decimal y multiplicando importes por 10/100.
    records = ws.get_all_records(value_render_option=ValueRenderOption.unformatted)
    return pd.DataFrame.from_records(records)


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner="Leyendo Google Sheets…")
def _parse_fecha_google(valor) -> pd.Timestamp:
    """Soporta tanto fechas que Google Sheets guardó como número de serie
    (epoch de Sheets: días desde 1899-12-30) cuando UNFORMATTED_VALUE trae un
    número, como fechas guardadas como texto ("31/05/2026")."""
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(valor, unit="D")
    return pd.to_datetime(valor, dayfirst=True, errors="coerce")


def load_movimientos() -> pd.DataFrame:
    if modo_demo_activo():
        return _demo_movimientos()
    client = get_client()
    df = _worksheet_to_df(client, "Movimientos")
    if df.empty:
        return df
    df["Fecha"] = df["Fecha"].apply(_parse_fecha_google)
    for col in ("Importe neto", "IVA"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["Importe total"] = df["Importe neto"] + df["IVA"]
    # signo: Gasto negativo, Cobro positivo — simplifica sumas de flujo de caja
    df["Importe firmado"] = df.apply(
        lambda r: r["Importe neto"] if r["Tipo"] == "Cobro" else -r["Importe neto"], axis=1
    )
    return df.dropna(subset=["Fecha"])


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def load_supuestos() -> dict:
    if modo_demo_activo():
        return _demo_supuestos()
    client = get_client()
    df = _worksheet_to_df(client, "Supuestos")
    return dict(zip(df["Parámetro"], df["Valor"]))


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def load_obligaciones_futuras() -> pd.DataFrame:
    if modo_demo_activo():
        return pd.DataFrame(columns=["Mes (AAAA-MM)", "Categoría", "Importe estimado", "Observación"])
    client = get_client()
    df = _worksheet_to_df(client, "Obligaciones_Futuras")
    if not df.empty:
        df["Importe estimado"] = pd.to_numeric(df["Importe estimado"], errors="coerce").fillna(0)
    return df


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def load_clientes() -> dict:
    """Cantidad de clientes por línea de negocio (hoja `Clientes`, agregada
    con extend_sheet_clientes.py). No es un dato transaccional, se actualiza
    cada tanto — se usa para Punto de Equilibrio y Sensibilidad."""
    if modo_demo_activo():
        return {"Liquidación de Sueldos": 48, "Tercerización de Estudios": 9,
                "Asesoramiento Laboral": 1, "Flujo Sole": 0, "Mandú": 6, "Otro ingreso": 0}
    client = get_client()
    try:
        df = _worksheet_to_df(client, "Clientes")
    except gspread.exceptions.WorksheetNotFound:
        return {}
    if df.empty:
        return {}
    df["Cantidad de clientes"] = pd.to_numeric(df["Cantidad de clientes"], errors="coerce").fillna(0)
    return dict(zip(df["Línea de negocio"], df["Cantidad de clientes"]))


# ──────────────────────────────────────────────────────────────────────────
# Datos de demostración (solo si no hay credenciales configuradas)
# ──────────────────────────────────────────────────────────────────────────


def _demo_movimientos() -> pd.DataFrame:
    import numpy as np

    rng = np.random.default_rng(42)
    hoy = pd.Timestamp.today().normalize()
    inicio = (hoy - pd.DateOffset(months=3)).replace(day=1)
    fechas = pd.date_range(inicio, hoy, freq="3D")

    filas = []
    for i, f in enumerate(fechas):
        cat_c = rng.choice(CATEGORIAS_COBRO, p=[0.35, 0.25, 0.05, 0.15, 0.15, 0.05])
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
    df = pd.DataFrame(filas)
    df["Importe total"] = df["Importe neto"] + df["IVA"]
    df["Importe firmado"] = df.apply(
        lambda r: r["Importe neto"] if r["Tipo"] == "Cobro" else -r["Importe neto"], axis=1
    )
    return df


def _demo_supuestos() -> dict:
    return {
        "Tipo de cambio USD/ARS": "1507",
        "Inflación mensual estimada": "0.019",
        "Objetivo margen neto de gestión": "0.30",
        "Objetivo caja SAC sin financiación — fecha": "2026-12-01",
        "Objetivo caja SAC sin financiación — umbral": "0",
        "TNA préstamo Banco Galicia": "0.32",
        "Monto préstamo Banco Galicia": "5086000",
        "Plazo préstamo Banco Galicia (meses)": "6",
    }


# ──────────────────────────────────────────────────────────────────────────
# Cálculos
# ──────────────────────────────────────────────────────────────────────────


def kpis_del_mes(df: pd.DataFrame, mes: pd.Period) -> dict:
    m = df[df["Fecha"].dt.to_period("M") == mes]
    ingresos = m.loc[m["Tipo"] == "Cobro", "Importe neto"].sum()
    egresos = m.loc[m["Tipo"] == "Gasto", "Importe neto"].sum()
    resultado = ingresos - egresos
    margen = (resultado / ingresos) if ingresos else 0.0
    return {
        "ingresos": ingresos, "egresos": egresos,
        "resultado": resultado, "margen": margen,
        "movimientos": len(m),
    }


def calcular_er(df: pd.DataFrame, mes: pd.Period) -> dict:
    """Reconstruye el Estado de Resultados del mes a partir de Movimientos,
    con el mismo criterio de líneas que la hoja `ER` del Excel original:
    separa costos directos y gastos de estructura de las partidas que NO son
    costo contable (Cuotas ARCA, Retiros de socios, Préstamo), que sí impactan
    el flujo de caja pero no el resultado de gestión."""
    m = df[df["Fecha"].dt.to_period("M") == mes]
    sumas = {b: 0.0 for b in ("ingresos", "costos_directos", "gastos_estructura", "impuestos", "no_operativo")}
    for bucket in sumas:
        cats = [c for c, b in CATEGORIA_A_LINEA_ER.items() if b == bucket]
        sumas[bucket] = m[m["Categoría"].isin(cats)]["Importe neto"].sum()

    ingresos = sumas["ingresos"]
    resultado_operativo_1 = ingresos - sumas["costos_directos"]
    resultado_operativo_2 = resultado_operativo_1 - sumas["gastos_estructura"]
    resultado_neto = resultado_operativo_2 - sumas["impuestos"]
    margen = (resultado_neto / ingresos) if ingresos else 0.0

    return {
        "ingresos": ingresos,
        "costos_directos": sumas["costos_directos"],
        "resultado_operativo_1": resultado_operativo_1,
        "gastos_estructura": sumas["gastos_estructura"],
        "resultado_operativo_2": resultado_operativo_2,
        "impuestos": sumas["impuestos"],
        "resultado_neto": resultado_neto,
        "margen": margen,
        "no_operativo": sumas["no_operativo"],
    }


LINEAS_DE_NEGOCIO = ["Liquidación de Sueldos", "Tercerización de Estudios", "Asesoramiento Laboral", "Flujo Sole", "Mandú"]


def participacion_por_linea(df: pd.DataFrame, mes: pd.Period) -> dict:
    """% que representa cada línea de negocio sobre el total de ingresos del
    mes — es la clave de reparto que usa el Excel original para asignar
    costos indirectos, impuestos y estructura a cada línea."""
    m = df[(df["Fecha"].dt.to_period("M") == mes) & (df["Tipo"] == "Cobro")]
    total = m["Importe neto"].sum()
    if not total:
        return {linea: 0.0 for linea in LINEAS_DE_NEGOCIO}
    return {
        linea: m[m["Categoría"] == linea]["Importe neto"].sum() / total
        for linea in LINEAS_DE_NEGOCIO
    }


def er_por_linea(df: pd.DataFrame, mes: pd.Period, supuestos: dict) -> pd.DataFrame:
    """Rentabilidad por línea de negocio, prorrateando costos directos,
    gastos de estructura e impuestos según la participación de cada línea
    en los ingresos del mes (mismo criterio que la hoja `ER x Línea` del
    Excel original). Como acá `Nómina` no está separada en directa/indirecta,
    todo el costo directo se prorratea junto — es una simplificación
    consciente frente al original, que si hace falta se puede afinar más
    adelante separando esa categoría."""
    er = calcular_er(df, mes)
    participacion = participacion_por_linea(df, mes)
    margen_objetivo = float(supuestos.get("Objetivo margen neto de gestión", 0.30) or 0.30)

    filas = []
    for linea in LINEAS_DE_NEGOCIO:
        p = participacion[linea]
        ingresos_linea = er["ingresos"] * p
        costos_directos_linea = er["costos_directos"] * p
        gastos_estructura_linea = er["gastos_estructura"] * p
        impuestos_linea = er["impuestos"] * p
        resultado_linea = ingresos_linea - costos_directos_linea - gastos_estructura_linea - impuestos_linea
        margen_linea = (resultado_linea / ingresos_linea) if ingresos_linea else 0.0
        filas.append({
            "Línea": linea, "Participación": p, "Ingresos": ingresos_linea,
            "Costos directos": -costos_directos_linea, "Gastos de estructura": -gastos_estructura_linea,
            "Impuestos": -impuestos_linea, "Resultado neto": resultado_linea,
            "Margen neto %": margen_linea, "Objetivo": margen_objetivo,
            "Cumple": margen_linea >= margen_objetivo,
        })
    return pd.DataFrame(filas)


def punto_equilibrio(df: pd.DataFrame, mes: pd.Period, clientes: dict) -> dict:
    """Estructura total a cubrir (costos directos + estructura + impuestos +
    Cuotas ARCA, esta última porque aunque no es gasto contable, sí es una
    salida de caja obligatoria) y cuántos clientes por línea hacen falta
    para cubrirla, usando el precio promedio actual por cliente de cada
    línea. Requiere la hoja `Clientes` — si está vacía, devuelve solo la
    parte en pesos (sin la lectura en cantidad de clientes)."""
    m = df[df["Fecha"].dt.to_period("M") == mes]
    er = calcular_er(df, mes)
    arca = m[(m["Tipo"] == "Gasto") & (m["Categoría"] == "Cuotas ARCA")]["Importe neto"].sum()

    estructura_operativa = er["costos_directos"] + er["gastos_estructura"] + arca
    estructura_total = estructura_operativa + er["impuestos"]
    margen_seguridad = er["ingresos"] - estructura_total
    margen_seguridad_pct = (margen_seguridad / er["ingresos"]) if er["ingresos"] else 0.0

    participacion = participacion_por_linea(df, mes)
    filas_lineas = []
    if clientes:
        for linea in LINEAS_DE_NEGOCIO:
            p = participacion[linea]
            ingresos_linea = er["ingresos"] * p
            clientes_actuales = clientes.get(linea, 0)
            precio_promedio = (ingresos_linea / clientes_actuales) if clientes_actuales else 0.0
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
    """Resultado de gestión mensual (antes de Impuesto a las Ganancias) si la
    cantidad de clientes variara `var_clientes` y el honorario promedio
    variara `var_honorarios` — replica la fórmula real de la hoja
    `Sensibilidad` del Excel original (leída de las celdas, no adivinada):

        Resultado = Ingresos_de_gestión × (1+honorarios) × (1+clientes)
                    × (1 − Ratio_facturado × Tasa_IIBB_efectiva)
                    − Costos_fijos_operativos

    Verificado contra las 25 celdas de la tabla original: coincide al
    centavo. Los ingresos de gestión (incluyen Mandú) escalan enteros con
    ambas variaciones — el 'Ratio facturado' es solo el % que
    representa la facturación operativa (sin Mandú) sobre el total, usado
    para estimar el efecto de IIBB sin tener que recalcularlo línea por
    línea. Los costos fijos (Nómina + Costos externos + Gastos de
    estructura) NO escalan con el escenario, igual que en el original."""
    er = calcular_er(df, mes)
    ingresos_base = er["ingresos"]
    costos_fijos_op = er["costos_directos"] + er["gastos_estructura"]

    m = df[(df["Fecha"].dt.to_period("M") == mes) & (df["Tipo"] == "Cobro")]
    facturacion_operativa = m[m["Categoría"] != "Mandú"]["Importe neto"].sum()
    ratio_facturado = (facturacion_operativa / ingresos_base) if ingresos_base else 0.0

    iibb_base = df[
        (df["Fecha"].dt.to_period("M") == mes) & (df["Tipo"] == "Gasto")
        & (df["Categoría"] == "Impuestos IIBB")
    ]["Importe neto"].sum()
    tasa_iibb_efectiva = (iibb_base / facturacion_operativa) if facturacion_operativa else 0.0

    ingresos_aj = ingresos_base * (1 + var_honorarios) * (1 + var_clientes) * (1 - ratio_facturado * tasa_iibb_efectiva)
    return ingresos_aj - costos_fijos_op


def matriz_sensibilidad(df: pd.DataFrame, mes: pd.Period,
                         pasos=(-0.2, -0.1, 0.0, 0.1, 0.2)) -> pd.DataFrame:
    """Filas = variación en cantidad de clientes (Q), columnas = variación
    en honorario promedio (P) — igual layout que 'Clientes ↓ / Honorarios →'
    del Excel original."""
    filas = []
    for vq in pasos:
        fila = {"Variación clientes": vq}
        for vp in pasos:
            fila[vp] = sensibilidad_resultado(df, mes, vq, vp)
        filas.append(fila)
    return pd.DataFrame(filas).set_index("Variación clientes")


def desglose_por_categoria(df: pd.DataFrame, mes: pd.Period, tipo: str) -> pd.DataFrame:
    m = df[(df["Fecha"].dt.to_period("M") == mes) & (df["Tipo"] == tipo)]
    return (
        m.groupby("Categoría")["Importe neto"].sum().sort_values(ascending=False).reset_index()
    )


def flujo_caja_real(df: pd.DataFrame) -> pd.DataFrame:
    """Evolución de la caja acumulada real, mes a mes, usando solo
    movimientos efectivamente cargados — sin proyección. Eso queda a cargo
    del informe mensual de Épsilon, no de este dashboard."""
    if df.empty:
        return pd.DataFrame(columns=["mes", "caja"])
    primer_mes = df["Fecha"].dt.to_period("M").min()
    ultimo_mes_con_datos = df["Fecha"].dt.to_period("M").max()
    meses = pd.period_range(primer_mes, ultimo_mes_con_datos, freq="M")

    filas = []
    caja = 0.0
    for m in meses:
        k = kpis_del_mes(df, m)
        caja += k["resultado"]
        filas.append({"mes": m.to_timestamp(), "caja": caja})
    return pd.DataFrame(filas)


def evaluar_semaforo(df: pd.DataFrame, supuestos: dict, mes: pd.Period) -> list:
    """Objetivos evaluados con datos reales del mes que se está mirando en
    el dashboard (no con la fecha de hoy del sistema)."""
    margen_obj = float(supuestos.get("Objetivo margen neto de gestión", 0.30) or 0.30)
    k = kpis_del_mes(df, mes)
    cumple_margen = k["margen"] >= margen_obj
    return [{
        "nombre": f"Margen neto del mes ≥ {margen_obj:.0%}",
        "valor": f"{k['margen']:.1%}", "cumple": cumple_margen,
    }]


# ──────────────────────────────────────────────────────────────────────────
# Interfaz
# ──────────────────────────────────────────────────────────────────────────

st.title("📊 Estudio Haberes — Dashboard financiero en tiempo real")

if modo_demo_activo():
    st.warning(
        "**MODO DEMO** — no se encontraron credenciales de Google Sheets en `st.secrets`. "
        "Mostrando datos sintéticos para probar la interfaz. Configurá "
        "`gcp_service_account` y `spreadsheet_id` según SETUP_GOOGLE_CLOUD.md "
        "para conectar la planilla real.",
        icon="⚠️",
    )

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

meses_disponibles = sorted(df_mov["Fecha"].dt.to_period("M").unique(), reverse=True)
mes_kpi = st.sidebar.selectbox("Mes de referencia para KPIs", meses_disponibles, index=0)

# ── KPIs del mes ───────────────────────────────────────────────────────
st.header(f"Indicadores — {mes_kpi}")
k = kpis_del_mes(df_mov, mes_kpi)
c1, c2, c3, c4 = st.columns(4)
c1.metric("Ingresos de gestión", f"${k['ingresos']:,.0f}")
c2.metric("Egresos totales", f"${k['egresos']:,.0f}")
c3.metric("Resultado neto", f"${k['resultado']:,.0f}")
c4.metric("Margen neto", f"{k['margen']:.1%}")

# ── Estado de Resultados en tiempo real ────────────────────────────────
st.subheader(f"Estado de Resultados — {mes_kpi}")
er = calcular_er(df_mov, mes_kpi)
filas_er = [
    ("Ingresos de gestión", er["ingresos"]),
    ("(−) Costos directos (Nómina + Costos externos + Otros)", -er["costos_directos"]),
    ("= Resultado operativo (antes de estructura)", er["resultado_operativo_1"]),
    ("(−) Gastos de estructura", -er["gastos_estructura"]),
    ("= Resultado operativo (antes de impuestos)", er["resultado_operativo_2"]),
    ("(−) Impuestos (IIBB + Ganancias)", -er["impuestos"]),
    ("= Resultado neto de gestión", er["resultado_neto"]),
]
df_er = pd.DataFrame(filas_er, columns=["Línea", "Importe"])
df_er["% s/ingresos"] = df_er["Importe"] / er["ingresos"] if er["ingresos"] else 0.0

col_tabla, col_donut = st.columns([3, 2])
with col_tabla:
    st.dataframe(
        df_er.style.format({"Importe": "${:,.0f}", "% s/ingresos": "{:.1%}"}),
        width="stretch", hide_index=True,
    )
    st.metric("Margen neto de gestión", f"{er['margen']:.1%}")
with col_donut:
    composicion = pd.DataFrame({
        "Componente": ["Costos directos", "Gastos de estructura", "Impuestos", "Resultado neto"],
        "Importe": [er["costos_directos"], er["gastos_estructura"], er["impuestos"],
                    max(er["resultado_neto"], 0)],
    })
    composicion = composicion[composicion["Importe"] > 0]
    fig_donut = go.Figure(data=[go.Pie(
        labels=composicion["Componente"], values=composicion["Importe"], hole=0.55,
        marker_colors=[ROJO_SELLO, DORADO_SELLO, "#8A8168", VERDE_LEDGER],
    )])
    plotly_layout_base(fig_donut, showlegend=True, legend=dict(orientation="h", y=-0.1))
    fig_donut.update_traces(textinfo="percent")
    st.plotly_chart(fig_donut, width="stretch")

if er["no_operativo"]:
    st.caption(
        f"Partidas fuera del resultado de gestión (Cuotas ARCA, Retiros de socios, "
        f"Préstamo): ${er['no_operativo']:,.0f}. Resultado neto de gestión "
        f"(${er['resultado_neto']:,.0f}) menos esas partidas = flujo de caja del mes "
        f"(${k['resultado']:,.0f}) — son dos números distintos a propósito: uno es "
        f"contable, el otro es caja."
    )

# ── ER x Línea de negocio ───────────────────────────────────────────────
st.subheader(f"Rentabilidad por línea de negocio — {mes_kpi}")
df_er_linea = er_por_linea(df_mov, mes_kpi, supuestos)
df_er_linea_mostrar = df_er_linea.copy()
df_er_linea_mostrar["Cumple"] = df_er_linea_mostrar["Cumple"].map({True: "✓", False: "✕"})
st.dataframe(
    df_er_linea_mostrar.style.format({
        "Participación": "{:.1%}", "Ingresos": "${:,.0f}", "Costos directos": "${:,.0f}",
        "Gastos de estructura": "${:,.0f}", "Impuestos": "${:,.0f}", "Resultado neto": "${:,.0f}",
        "Margen neto %": "{:.1%}", "Objetivo": "{:.0%}",
    }),
    width="stretch", hide_index=True,
)
fig_margen_linea = go.Figure()
colores_margen = [VERDE_LEDGER if c else ROJO_SELLO for c in df_er_linea["Cumple"]]
fig_margen_linea.add_trace(go.Bar(
    y=df_er_linea["Línea"], x=df_er_linea["Margen neto %"], orientation="h",
    marker_color=colores_margen, text=[f"{v:.1%}" for v in df_er_linea["Margen neto %"]],
    textposition="outside", name="Margen real",
))
objetivo_linea = df_er_linea["Objetivo"].iloc[0] if not df_er_linea.empty else 0
fig_margen_linea.add_vline(x=objetivo_linea, line_dash="dash", line_color=DORADO_SELLO,
                            annotation_text=f"Objetivo {objetivo_linea:.0%}")
plotly_layout_base(fig_margen_linea, xaxis_tickformat=".0%", showlegend=False,
                    xaxis_title="Margen neto %", yaxis_title=None)
st.plotly_chart(fig_margen_linea, width="stretch")
st.caption(
    "Los costos se reparten entre líneas según su participación en los ingresos del "
    "mes (mismo criterio que el Excel original). Como acá 'Nómina' es una sola "
    "categoría (no separada en directa/indirecta), este reparto es una aproximación "
    "algo más gruesa que la del Excel — si en algún momento hace falta más precisión, "
    "se puede separar esa categoría más adelante."
)

# ── Punto de equilibrio ─────────────────────────────────────────────────
st.subheader(f"Punto de equilibrio y margen de seguridad — {mes_kpi}")
clientes = load_clientes()
peq = punto_equilibrio(df_mov, mes_kpi, clientes)
c1, c2, c3 = st.columns(3)
c1.metric("Estructura total a cubrir", f"${peq['estructura_total']:,.0f}")
c2.metric("Margen de seguridad", f"${peq['margen_seguridad']:,.0f}")
c3.metric("Margen de seguridad %", f"{peq['margen_seguridad_pct']:.1%}")
if not peq["por_linea"].empty:
    st.dataframe(
        peq["por_linea"].style.format({
            "Precio promedio/cliente": "${:,.0f}",
            "Clientes necesarios (si cubriera toda la estructura sola)": "{:,.1f}",
        }),
        width="stretch", hide_index=True,
    )
else:
    st.info(
        "No hay datos en la hoja 'Clientes' todavía — corré extend_sheet_clientes.py "
        "para agregarla y ver el punto de equilibrio en cantidad de clientes por línea."
    )

# ── Semáforo de objetivos ──────────────────────────────────────────────
st.subheader("Semáforo de objetivos")
objetivos = evaluar_semaforo(df_mov, supuestos, mes_kpi)
for o in objetivos:
    clase = "semaforo-ok" if o["cumple"] else "semaforo-no"
    icono = "✓" if o["cumple"] else "✕"
    st.markdown(
        f'<span class="semaforo-badge {clase}">{icono} {o["nombre"]}</span> '
        f'<span style="font-family:IBM Plex Mono, monospace;">{o["valor"]}</span>',
        unsafe_allow_html=True,
    )

# ── Flujo de caja acumulado (real) ──────────────────────────────────────
st.subheader("Flujo de caja acumulado — real")
serie_caja = flujo_caja_real(df_mov)
fig_caja = go.Figure()
fig_caja.add_trace(go.Scatter(x=serie_caja["mes"], y=serie_caja["caja"], mode="lines+markers",
                               name="Real", line=dict(color=INK, width=2),
                               fill="tozeroy", fillcolor=f"{INK}11"))
plotly_layout_base(fig_caja, showlegend=False, yaxis_title="Caja acumulada ($)")
st.plotly_chart(fig_caja, width="stretch")
st.caption(
    "Evolución de la caja acumulada con datos reales cargados hasta la fecha. La "
    "proyección a futuro queda a cargo del informe mensual de Épsilon."
)

# ── Análisis de sensibilidad ────────────────────────────────────────────
st.subheader(f"Análisis de sensibilidad — {mes_kpi}")
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
    text=[[f"${v:,.0f}" for v in fila] for fila in matriz.values],
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
    "Resultado de gestión mensual (antes de Impuesto a las Ganancias) estimado según "
    "cómo varíen la cantidad de clientes (Q) y el honorario promedio (P), manteniendo "
    "fijos los costos operativos — misma lógica y mismos números que la tabla de "
    "sensibilidad del Excel original."
)

# ── Desglose por categoría ─────────────────────────────────────────────
col_a, col_b = st.columns(2)
with col_a:
    st.subheader("Ingresos por línea de negocio")
    df_ing = desglose_por_categoria(df_mov, mes_kpi, "Cobro")
    fig_ing = go.Figure(go.Bar(x=df_ing["Importe neto"], y=df_ing["Categoría"], orientation="h",
                                marker_color=VERDE_LEDGER))
    plotly_layout_base(fig_ing, yaxis=dict(autorange="reversed"), xaxis_title="$")
    st.plotly_chart(fig_ing, width="stretch")
with col_b:
    st.subheader("Egresos por categoría")
    df_gas = desglose_por_categoria(df_mov, mes_kpi, "Gasto")
    fig_gas = go.Figure(go.Bar(x=df_gas["Importe neto"], y=df_gas["Categoría"], orientation="h",
                                marker_color=ROJO_SELLO))
    plotly_layout_base(fig_gas, yaxis=dict(autorange="reversed"), xaxis_title="$")
    st.plotly_chart(fig_gas, width="stretch")

# ── Tabla de movimientos filtrada ──────────────────────────────────────
st.subheader("Movimientos (según filtros)")
st.dataframe(df_filtrado.sort_values("Fecha", ascending=False), width="stretch")
st.download_button(
    "Descargar CSV filtrado", df_filtrado.to_csv(index=False).encode("utf-8"),
    file_name="movimientos_filtrados.csv", mime="text/csv",
)