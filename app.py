"""
SIDOP – Sistema Inteligente de Diagnóstico Operacional de Pozos
Quintana Energy
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────

GDRIVE_FILE_ID = "1bur7CrEDEy5BjC4dV-FwK2Bi49XPt3yu"
EXCEL_PATH = Path("/tmp/Diagnostico_Pozos_PowerBI.xlsx")

def descargar_excel():
    """Descarga el Excel desde Google Drive si no existe localmente."""
    if not EXCEL_PATH.exists():
        try:
            import gdown
            url = f"https://drive.google.com/uc?id={GDRIVE_FILE_ID}"
            gdown.download(url, str(EXCEL_PATH), quiet=True)
        except Exception as e:
            st.error(f"No se pudo descargar el archivo de datos: {e}")
            st.stop()

descargar_excel()

# Quintana Energy — Brand Colors
C_DARK   = "#33492D"   # Verde Bosque  → header
C_PRIM   = "#6B7B38"   # Verde Oliva   → accents, labels
C_NAVY   = "#1B4B6C"   # Azul Marino   → charts, secondary
C_BG     = "#DAE0E5"   # Gris Azulado  → card backgrounds
C_CREAM  = "#FFEAC6"   # Crema Cálido  → recomendación
C_LIME   = "#E2FF87"   # Verde Lima    → prod oil
C_RED    = "#C0392B"   # Alerta roja
C_WHITE  = "#FFFFFF"
C_TEXT   = "#1a1a1a"

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="SIDOP – Quintana Energy",
    page_icon="🛢️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&family=Inter:wght@400;500;600&display=swap');

html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
    color: {C_TEXT};
}}

/* ── HEADER ── */
.qe-header {{
    background: {C_DARK};
    padding: 16px 24px 14px;
    border-radius: 8px;
    margin-bottom: 20px;
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
}}
.qe-header h1 {{
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-size: 1.4rem;
    font-weight: 800;
    margin: 0;
    letter-spacing: .3px;
    text-transform: uppercase;
}}
.qe-header .subtitle {{
    color: #a8c89a;
    font-size: 0.76rem;
    margin-top: 5px;
}}
.qe-header .brand {{
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    font-size: 1rem;
    opacity: .9;
}}

/* ── KPI CARD GRANDE ── */
.kpi-card {{
    background: {C_BG};
    border-radius: 10px;
    padding: 20px 16px 16px;
    text-align: center;
}}
.kpi-number {{
    font-family: 'Montserrat', sans-serif;
    font-size: 2.4rem;
    font-weight: 800;
    color: {C_NAVY};
    line-height: 1;
}}
.kpi-label {{
    font-size: 0.77rem;
    color: #555;
    margin-top: 6px;
    line-height: 1.4;
    font-weight: 500;
}}

/* ── KPI PEQUEÑO (detalle) ── */
.kpi-sm {{
    background: {C_BG};
    border-radius: 8px;
    padding: 12px 8px 10px;
    text-align: center;
}}
.kpi-sm-label {{
    font-size: 0.67rem;
    color: #555;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: .4px;
}}
.kpi-sm-value {{
    font-family: 'Montserrat', sans-serif;
    font-size: 1.45rem;
    font-weight: 700;
    color: {C_PRIM};
    margin-top: 3px;
}}

/* ── INFO CARD (header detalle) ── */
.info-card {{
    background: {C_BG};
    border-radius: 8px;
    padding: 12px 14px;
    height: 100%;
}}
.info-card-label {{
    font-size: 0.68rem;
    color: {C_PRIM};
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .5px;
}}
.info-card-value {{
    font-family: 'Montserrat', sans-serif;
    font-size: 1.1rem;
    font-weight: 600;
    color: {C_TEXT};
    margin-top: 5px;
    line-height: 1.2;
}}

/* ── PROD CARD ── */
.prod-card {{
    background: white;
    border: 1.5px solid {C_BG};
    border-radius: 8px;
    padding: 12px 14px;
}}
.prod-card-label {{
    font-size: 0.68rem;
    color: {C_NAVY};
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .5px;
}}
.prod-card-value {{
    font-family: 'Montserrat', sans-serif;
    font-size: 1.6rem;
    font-weight: 700;
    color: {C_PRIM};
    margin-top: 4px;
}}

/* ── ALERTA CARD ── */
.alerta-card {{
    background: #fdf0f0;
    border: 1.5px solid {C_RED};
    border-radius: 8px;
    padding: 12px 14px;
    height: 100%;
}}
.alerta-card-label {{
    font-size: 0.68rem;
    color: {C_RED};
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .5px;
}}
.alerta-card-value {{
    font-size: 0.82rem;
    color: {C_RED};
    margin-top: 6px;
    line-height: 1.5;
    font-weight: 500;
}}

.alerta-ok {{
    background: #f0f7f0;
    border: 1.5px solid {C_PRIM};
    border-radius: 8px;
    padding: 12px 14px;
    height: 100%;
}}
.alerta-ok-label {{
    font-size: 0.68rem;
    color: {C_PRIM};
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .5px;
}}
.alerta-ok-value {{
    font-size: 0.82rem;
    color: {C_PRIM};
    margin-top: 6px;
    font-weight: 500;
}}

/* ── DIAGNÓSTICO OPERATIVO ── */
.diag-card {{
    background: #EEF3FB;
    border-left: 4px solid {C_NAVY};
    border-radius: 6px;
    padding: 14px 18px;
    margin-top: 6px;
}}
.diag-card-header {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 8px;
}}
.diag-card-label {{
    font-size: 0.72rem;
    color: {C_NAVY};
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .5px;
}}
.diag-card-fecha {{
    font-size: 0.72rem;
    color: #666;
    font-weight: 500;
}}
.diag-subcategoria {{
    display: inline-block;
    background: {C_NAVY};
    color: white;
    font-size: 0.68rem;
    font-weight: 600;
    padding: 2px 10px;
    border-radius: 20px;
    margin-bottom: 8px;
}}
.diag-card-text {{
    font-size: 0.88rem;
    color: {C_TEXT};
    line-height: 1.6;
    font-weight: 500;
}}
.diag-observacion {{
    font-size: 0.83rem;
    color: #555;
    line-height: 1.5;
    margin-top: 8px;
    border-top: 1px solid #ccd9ee;
    padding-top: 8px;
}}

/* ── RECOMENDACIÓN ── */
.rec-card {{
    background: {C_CREAM};
    border-left: 4px solid {C_PRIM};
    border-radius: 6px;
    padding: 14px 18px;
    margin-top: 6px;
}}
.rec-card-label {{
    font-size: 0.72rem;
    color: {C_PRIM};
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .5px;
    margin-bottom: 8px;
}}
.rec-card-text {{
    font-size: 0.88rem;
    color: {C_TEXT};
    line-height: 1.6;
}}

/* ── SECTION TITLE ── */
.section-title {{
    background: {C_PRIM};
    color: white;
    padding: 6px 14px;
    border-radius: 4px;
    font-family: 'Montserrat', sans-serif;
    font-size: 0.8rem;
    font-weight: 700;
    letter-spacing: .5px;
    text-transform: uppercase;
    margin: 18px 0 10px;
    display: inline-block;
}}

/* ── SIDEBAR ── */
section[data-testid="stSidebar"] > div {{
    background: #f4f6f8;
}}

/* Spacing */
.block-container {{
    padding-top: 1rem;
    padding-bottom: 2rem;
    max-width: 1400px;
}}

/* ── Header de Streamlit: solo el botón del sidebar, sin fondo visible ── */
header[data-testid="stHeader"] {{
    background: transparent !important;
    border-bottom: none !important;
    box-shadow: none !important;
}}

/* ── TABS: sticky dentro del contenedor de scroll ── */
div[data-testid="stTabs"] button {{
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    padding: 10px 24px !important;
}}
div[data-testid="stTabs"] > div:first-child {{
    position: sticky;
    top: 0;
    z-index: 999;
    background: white;
    padding: 6px 0 4px;
    border-bottom: 3px solid {C_PRIM};
    box-shadow: 0 4px 10px rgba(0,0,0,0.12);
}}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def safe_val(val, decimals=1, suffix=""):
    """Formatea un valor numérico de forma segura."""
    try:
        v = float(val)
        if np.isnan(v):
            return "–"
        if decimals == 0:
            return f"{int(round(v))}{suffix}"
        return f"{v:.{decimals}f}{suffix}"
    except:
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return "–"
        return f"{val}{suffix}"


def safe_date(val):
    """Formatea una fecha de forma segura."""
    if val is None:
        return "–"
    try:
        dt = pd.to_datetime(val)
        if pd.isna(dt):
            return "–"
        return dt.strftime("%d/%m/%Y")
    except:
        return str(val)[:10] if val else "–"


def color_riesgo(val):
    """Estilo CSS por nivel de riesgo."""
    estilos = {
        "Critico": "background-color:#f8d7da; color:#721c24; font-weight:600;",
        "Alto":    "background-color:#fff3cd; color:#856404; font-weight:600;",
        "Medio":   "background-color:#cce5ff; color:#004085; font-weight:600;",
        "Bajo":    "background-color:#d4edda; color:#155724; font-weight:600;",
    }
    return estilos.get(str(val), "")


# ─────────────────────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────────────────────

@st.cache_data(ttl=None, show_spinner="Cargando datos desde Excel...")
def load_all_data(path: str):
    xl = pd.ExcelFile(path)
    return {sheet: xl.parse(sheet) for sheet in xl.sheet_names}


@st.cache_data(ttl=None, show_spinner="Indexando y procesando datos...")
def preindexar_por_pozo(path: str):
    """Pre-agrupa y pre-procesa todas las hojas por NOMBRE_POZO.
    Fechas, tipos y ordenamiento se hacen UNA SOLA VEZ aquí dentro del cache.
    """
    datos = load_all_data(path)

    def agrupar(df, col="NOMBRE_POZO"):
        if df.empty or col not in df.columns:
            return {}
        # Normalizar a MAYÚSCULAS para que el lookup sea case-insensitive
        df = df.copy()
        df[col] = df[col].astype(str).str.strip().str.upper()
        return {k: v.reset_index(drop=True) for k, v in df.groupby(col)}

    # ── Tendencia Producción: fechas + numéricos + sort ──
    raw_prod = datos.get("Tendencia_Produccion", pd.DataFrame()).copy()
    if not raw_prod.empty and "FECHA" in raw_prod.columns:
        raw_prod["FECHA"] = pd.to_datetime(raw_prod["FECHA"], errors="coerce")
        for c in ["TOTAL_LIQUIDO", "PROD_OIL"]:
            if c in raw_prod.columns:
                raw_prod[c] = pd.to_numeric(raw_prod[c], errors="coerce")
        raw_prod = raw_prod.sort_values("FECHA").dropna(subset=["FECHA"])

    # ── Tendencia Niveles: fechas + numéricos + sort ──
    raw_niv = datos.get("Tendencia_Niveles", pd.DataFrame()).copy()
    if not raw_niv.empty and "FECHA_NVL" in raw_niv.columns:
        raw_niv["FECHA_NVL"] = pd.to_datetime(raw_niv["FECHA_NVL"], errors="coerce")
        if "SUMERGENCIA" in raw_niv.columns:
            raw_niv["SUMERGENCIA"] = pd.to_numeric(raw_niv["SUMERGENCIA"], errors="coerce")
        raw_niv = raw_niv.sort_values("FECHA_NVL").dropna(subset=["FECHA_NVL", "SUMERGENCIA"])

    # ── Tendencia Dinamometría: fechas + numéricos + sort ──
    raw_dina = datos.get("Tendencia_Dinamometria", pd.DataFrame()).copy()
    if not raw_dina.empty and "FECHA_DINA" in raw_dina.columns:
        raw_dina["FECHA_DINA"] = pd.to_datetime(raw_dina["FECHA_DINA"], errors="coerce")
        for c in ["BBA_LLENADO_DE_BOMBA", "BBA_EFICIENCIA_VOLUMETRICA"]:
            if c in raw_dina.columns:
                raw_dina[c] = pd.to_numeric(raw_dina[c], errors="coerce")
        raw_dina = raw_dina.sort_values("FECHA_DINA").dropna(subset=["FECHA_DINA"])

    # ── Eventos Riesgo: fechas + sort ──
    raw_ev = datos.get("Eventos_Riesgo", pd.DataFrame()).copy()
    if not raw_ev.empty:
        fecha_col_ev = next((c for c in ["FECHA_INICIO", "FECHA_EVENTO"] if c in raw_ev.columns), None)
        if fecha_col_ev:
            raw_ev[fecha_col_ev] = pd.to_datetime(raw_ev[fecha_col_ev], errors="coerce")
            raw_ev = raw_ev.sort_values(fecha_col_ev, ascending=False)

    # ── Cartas: todas (CF y CS) ──
    raw_cartas = datos.get("Dinas_Cartas_Ultimas3", pd.DataFrame()).copy()

    return {
        "prod":    agrupar(raw_prod),
        "niv":     agrupar(raw_niv),
        "dina":    agrupar(raw_dina),
        "eventos": agrupar(raw_ev),
        "cartas":  agrupar(raw_cartas),
        "diag":    agrupar(datos.get("Diagnostico_Cartas",      pd.DataFrame())),
        "diag_op": agrupar(datos.get("Historico_Diagnosticos",  pd.DataFrame())),
    }


@st.cache_data(ttl=None, show_spinner=False)
def get_ult_diagnostico(path: str) -> pd.DataFrame:
    """
    Retorna el diagnóstico más reciente por pozo desde Historico_Diagnosticos.
    Funciona aunque en el futuro haya múltiples entradas por pozo.
    """
    datos = load_all_data(path)
    hd = datos.get("Historico_Diagnosticos", pd.DataFrame()).copy()
    if hd.empty:
        return pd.DataFrame()
    # Normalizar nombre de columna de pozo
    col_pozo = next((c for c in ["NOMBRE_POZO", "NOMBRE_POZ"] if c in hd.columns), None)
    if col_pozo is None:
        return pd.DataFrame()
    if col_pozo != "NOMBRE_POZO":
        hd = hd.rename(columns={col_pozo: "NOMBRE_POZO"})
    # Ordenar por fecha descendente y tomar el más reciente por pozo
    if "DIAG_FECHA" in hd.columns:
        hd["DIAG_FECHA"] = pd.to_datetime(hd["DIAG_FECHA"], errors="coerce")
        hd = hd.sort_values("DIAG_FECHA", ascending=False)
    return hd.groupby("NOMBRE_POZO", as_index=False).first()


try:
    datos = load_all_data(str(EXCEL_PATH))
except FileNotFoundError:
    st.error(
        f"⚠️ No se encontró el archivo:\n`{EXCEL_PATH}`\n\n"
        "Verificá que el archivo existe y la ruta es correcta."
    )
    st.stop()
except Exception as e:
    st.error(f"Error al abrir el archivo Excel: {e}")
    st.stop()

# Extraer hojas
maestro   = datos.get("Maestro_Pozos",        pd.DataFrame())
tend_niv  = datos.get("Tendencia_Niveles",     pd.DataFrame())
tend_prod = datos.get("Tendencia_Produccion",  pd.DataFrame())
tend_dina = datos.get("Tendencia_Dinamometria",pd.DataFrame())
eventos   = datos.get("Eventos_Riesgo",        pd.DataFrame())
cartas    = datos.get("Dinas_Cartas_Ultimas3", pd.DataFrame())
diag_cart = datos.get("Diagnostico_Cartas",    pd.DataFrame())
acciones  = datos.get("Acciones",              pd.DataFrame())

# Índice por pozo — lookup instantáneo al cambiar de pozo
idx = preindexar_por_pozo(str(EXCEL_PATH))

# Último diagnóstico por pozo — fuente para el análisis de sumergencia en Tab 1
ult_diag_todos = get_ult_diagnostico(str(EXCEL_PATH))

# Timestamp de actualización
try:
    mtime = EXCEL_PATH.stat().st_mtime
    last_update = datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M")
except:
    last_update = "–"

# Normalizar numéricos en maestro
NUM_COLS = [
    "SUMERGENCIA", "BBA_LLENADO_DE_BOMBA", "BBA_EFICIENCIA_VOLUMETRICA",
    "AIBRE_SOLICITACION_DE_ESTRUCT", "AIBRR_PORCENTAJE", "AIBEB_PORCENTAJE",
    "PROD_OIL", "TOTAL_LIQUIDO", "PROD_GAS",
    "EVENTOS_12M", "RIESGO_HISTORICO", "RIESGO_OPERACIONAL_SCORE",
    "AIB_CARRERA_DEL_VASTAGO", "MOTOR_DIAMETRO_POLEA",
    "PROF_ENT_BOMBA", "REGIMEN_PRIM", "MOTOR_RPM",
]
for col in NUM_COLS:
    if col in maestro.columns:
        maestro[col] = pd.to_numeric(maestro[col], errors="coerce")

# ─────────────────────────────────────────────────────────────
# SIDEBAR — FILTROS GLOBALES
# ─────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown(f"""
    <div style="background:{C_DARK}; padding:14px 14px 12px; border-radius:8px; margin-bottom:18px;">
      <div style="color:white; font-family:Montserrat,sans-serif; font-weight:800;
                  font-size:1.1rem; letter-spacing:.3px;">SIDOP</div>
      <div style="color:#a8c89a; font-size:0.72rem; margin-top:3px;">
        Sistema de Diagnóstico Operacional<br>Quintana Energy
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### Filtros")

    zonas_opts = sorted(maestro["ZONA"].dropna().unique()) if "ZONA" in maestro.columns else []
    sel_zona = st.multiselect("Zona", options=zonas_opts, placeholder="Todas")

    bat_opts = sorted(maestro["BATERIA_FINAL"].dropna().unique()) if "BATERIA_FINAL" in maestro.columns else []
    sel_bat = st.multiselect("Batería", options=bat_opts, placeholder="Todas")

    metodo_opts = sorted(maestro["SISTEMA_EXTRACTIVO"].dropna().unique()) if "SISTEMA_EXTRACTIVO" in maestro.columns else []
    sel_metodo = st.multiselect("Sistema extractivo", options=metodo_opts, placeholder="Todos")

    estado_op_opts = sorted(maestro["ESTADO_OPERATIVO"].dropna().unique()) if "ESTADO_OPERATIVO" in maestro.columns else []
    sel_estado_op = st.multiselect("Estado operativo", options=estado_op_opts, placeholder="Todos")

    st.markdown("---")
    if st.button("🔄 Actualizar datos", use_container_width=True, type="primary"):
        if EXCEL_PATH.exists():
            EXCEL_PATH.unlink()   # borra local → fuerza re-descarga desde Drive
        st.cache_data.clear()
        st.rerun()
    st.caption(f"Última actualización: {last_update}")

# ─────────────────────────────────────────────────────────────
# APLICAR FILTROS
# ─────────────────────────────────────────────────────────────

df = maestro.copy()
if sel_zona:
    df = df[df["ZONA"].isin(sel_zona)]
if sel_bat:
    df = df[df["BATERIA_FINAL"].isin(sel_bat)]
if sel_metodo:
    df = df[df["SISTEMA_EXTRACTIVO"].isin(sel_metodo)]
if sel_estado_op:
    df = df[df["ESTADO_OPERATIVO"].isin(sel_estado_op)]

# ─────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────

tab1, tab2 = st.tabs(["📊   Resumen Operativo", "🔍   Detalle de Pozo"])


# ══════════════════════════════════════════════════════════════
#  TAB 1  —  RESUMEN OPERATIVO
# ══════════════════════════════════════════════════════════════

with tab1:

    # df_tab1: solo pozos EN EXTRACCION según Cierre Diario, excluye Pistoneo/Surgente
    EXCLUIR_SISTEMAS = {"PISTONEO", "SURGENTE"}
    if "EN_EXTRACCION_CIERRE" in df.columns:
        df_tab1 = df[df["EN_EXTRACCION_CIERRE"] == True].copy()
        if "SISTEMA_EXTRACTIVO" in df_tab1.columns:
            df_tab1 = df_tab1[~df_tab1["SISTEMA_EXTRACTIVO"].isin(EXCLUIR_SISTEMAS)]
    else:
        df_tab1 = df[~df["SISTEMA_EXTRACTIVO"].isin(EXCLUIR_SISTEMAS)].copy() \
            if "SISTEMA_EXTRACTIVO" in df.columns else df.copy()

    # ── Último diagnóstico — fuente completa de Historico_Diagnosticos ──
    # NO se filtra por Cierre Diario: el universo lo define la hoja de diagnósticos.
    if not ult_diag_todos.empty and "NOMBRE_POZO" in ult_diag_todos.columns:
        ult_diag_t1 = ult_diag_todos.copy()
        # Calcular ZONA desde Diccionario_Zonas (BATERIA → ZONA)
        _dz = datos.get("Diccionario_Zonas", pd.DataFrame())
        if not _dz.empty and "BATERIA" in _dz.columns and "ZONA" in _dz.columns:
            _mapa_z = dict(zip(_dz["BATERIA"].astype(str), _dz["ZONA"]))
            if "BATERIA" in ult_diag_t1.columns:
                ult_diag_t1["ZONA"] = ult_diag_t1["BATERIA"].astype(str).map(_mapa_z).fillna("SIN_ZONA")
        if "ZONA" not in ult_diag_t1.columns:
            # Fallback: inferir desde nombre del pozo
            ult_diag_t1["ZONA"] = ult_diag_t1["NOMBRE_POZO"].apply(
                lambda n: "CL" if str(n).upper().startswith("CL-")
                          else ("CS" if str(n).upper().startswith("CS-") else "SIN_ZONA")
            )
    else:
        ult_diag_t1 = pd.DataFrame()

    # ── Aplicar filtros del sidebar a ult_diag_t1 ──
    if not ult_diag_t1.empty:
        # Zona y Batería están disponibles en ult_diag_t1
        if sel_zona and "ZONA" in ult_diag_t1.columns:
            ult_diag_t1 = ult_diag_t1[ult_diag_t1["ZONA"].isin(sel_zona)]
        if sel_bat and "BATERIA" in ult_diag_t1.columns:
            ult_diag_t1 = ult_diag_t1[
                ult_diag_t1["BATERIA"].astype(str).isin([str(b) for b in sel_bat])
            ]
        # Sistema extractivo y Estado operativo: filtrar por NOMBRE_POZO del df ya filtrado
        if sel_metodo or sel_estado_op:
            _pozos_gate = set(df["NOMBRE_POZO"].astype(str).str.strip().str.upper())
            ult_diag_t1 = ult_diag_t1[
                ult_diag_t1["NOMBRE_POZO"].astype(str).str.strip().str.upper().isin(_pozos_gate)
            ]

    _usa_diag = not ult_diag_t1.empty and "DIAG_ESTADO_OP" in ult_diag_t1.columns

    # ── Header ──
    st.markdown(f"""
    <div class="qe-header">
      <div>
        <h1>Diagnóstico Operacional de Pozos</h1>
        <div class="subtitle">Última actualización datos: {last_update}</div>
      </div>
      <div class="brand">Quintana Energy</div>
    </div>
    """, unsafe_allow_html=True)

    # ── KPIs — universo: Historico_Diagnosticos (último diagnóstico por pozo) ──
    if _usa_diag:
        total_pozos   = len(ult_diag_t1)
        # Pozos con sub-categoría asignada (col F — DIAG_SUBCATEGORIA no vacía)
        if "DIAG_SUBCATEGORIA" in ult_diag_t1.columns:
            _mask_sub_asig = (
                ult_diag_t1["DIAG_SUBCATEGORIA"].notna()
                & (ult_diag_t1["DIAG_SUBCATEGORIA"].astype(str).str.strip() != "")
                & (ult_diag_t1["DIAG_SUBCATEGORIA"].astype(str).str.upper() != "NAN")
            )
            total_con_diag = int(_mask_sub_asig.sum())
        else:
            total_con_diag = total_pozos
        en_potencial  = int((ult_diag_t1["DIAG_ESTADO_OP"] == "En potencial").sum())
        subexplotados = int(ult_diag_t1["DIAG_ESTADO_OP"].str.contains("Subexplotado", na=False).sum())
        # 4to KPI: pozos clasificados como Parado en la sub-categoría
        parados = int(
            (ult_diag_t1["DIAG_SUBCATEGORIA"].astype(str).str.strip().str.upper() == "PARADO").sum()
        ) if "DIAG_SUBCATEGORIA" in ult_diag_t1.columns else 0
    else:
        # fallback: fuente Maestro (ESTADO_OPERATIVO calculado desde niveles)
        total_pozos   = len(df_tab1)
        total_con_diag = total_pozos
        en_potencial  = len(df_tab1[df_tab1["ESTADO_OPERATIVO"] == "En potencial"]) if "ESTADO_OPERATIVO" in df_tab1.columns else 0
        subexplotados = len(df_tab1[df_tab1["ESTADO_OPERATIVO"].str.contains("Subexplotado", na=False)]) if "ESTADO_OPERATIVO" in df_tab1.columns else 0
        parados       = 0

    k1, k2, k3, k4, k5 = st.columns(5)
    for col_w, num, label in [
        (k1, total_pozos,    "Total"),
        (k2, total_con_diag, "Total<br>con Diagnóstico"),
        (k3, en_potencial,   "Pozos<br>En Potencial"),
        (k4, subexplotados,  "Pozos<br>Sub-explotados"),
        (k5, parados,        "Pozos<br>Parados"),
    ]:
        with col_w:
            st.markdown(f"""
            <div class="kpi-card">
              <div class="kpi-number">{num}</div>
              <div class="kpi-label">{label}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:18px'></div>", unsafe_allow_html=True)

    # ── Gráficos ──
    g1, g2 = st.columns(2)

    with g1:
        st.markdown("**Total Pozos por Zona**")
        _zona_src = ult_diag_t1 if (_usa_diag and "ZONA" in ult_diag_t1.columns) else df_tab1
        if "ZONA" in _zona_src.columns and len(_zona_src) > 0:
            zona_cnt = (
                _zona_src.groupby("ZONA", as_index=False)
                .size()
                .rename(columns={"size": "Total"})
                .sort_values("Total", ascending=False)
            )
            fig_zona = px.bar(
                zona_cnt, x="ZONA", y="Total",
                text="Total",
                color_discrete_sequence=[C_NAVY],
            )
            fig_zona.update_traces(textposition="outside", textfont_size=13, cliponaxis=False)
            fig_zona.update_layout(
                plot_bgcolor="white", paper_bgcolor="white",
                height=300, margin=dict(l=10, r=10, t=50, b=30),
                xaxis_title="Zona", yaxis_title="Total Pozos",
                font=dict(family="Inter", size=12),
                showlegend=False,
            )
            st.plotly_chart(fig_zona, use_container_width=True)
        else:
            st.info("Sin datos de zona.")

    with g2:
        st.markdown("**Estado Operativo — Diagnóstico de Sumergencia**")
        if _usa_diag:
            est_src = ult_diag_t1.rename(columns={"DIAG_ESTADO_OP": "ESTADO_OPERATIVO"})
        elif "ESTADO_OPERATIVO" in df_tab1.columns:
            est_src = df_tab1
        else:
            est_src = pd.DataFrame()

        if not est_src.empty and "ESTADO_OPERATIVO" in est_src.columns:
            est_cnt = (
                est_src.groupby("ESTADO_OPERATIVO", as_index=False)
                .size()
                .rename(columns={"size": "Total"})
                .sort_values("Total", ascending=False)
            )
            fig_est = px.bar(
                est_cnt, x="ESTADO_OPERATIVO", y="Total",
                text="Total",
                color_discrete_sequence=[C_NAVY],
            )
            fig_est.update_traces(textposition="outside", textfont_size=13, cliponaxis=False)
            fig_est.update_layout(
                plot_bgcolor="white", paper_bgcolor="white",
                height=320, margin=dict(l=10, r=10, t=50, b=90),
                xaxis_title="Estado Operativo", yaxis_title="Total Pozos",
                font=dict(family="Inter", size=11),
                showlegend=False,
                xaxis_tickangle=-20,
            )
            st.plotly_chart(fig_est, use_container_width=True)
        else:
            st.info("Sin datos de estado operativo.")

    # ── Gráficos Sub-categoría — fuente: Historico_Diagnosticos ──
    rangos = [
        ("Subexplotado bajo entre 201-400m",      C_PRIM),
        ("Subexplotado moderado entre 401-600m",   "#E67E22"),
        ("Subexplotado alto > 600m",               C_RED),
    ]

    # Elegir fuente: diagnóstico experto (preferido) o fallback en Maestro
    if _usa_diag and "DIAG_SUBCATEGORIA" in ult_diag_t1.columns:
        _src_sub = ult_diag_t1.copy()
        _col_estado = "DIAG_ESTADO_OP"
    elif "DIAG_SUBCATEGORIA" in df_tab1.columns and "ESTADO_OPERATIVO" in df_tab1.columns:
        _src_sub = df_tab1.copy()
        _col_estado = "ESTADO_OPERATIVO"
    else:
        _src_sub = pd.DataFrame()
        _col_estado = None

    if not _src_sub.empty and _col_estado:
        tiene_datos = any(
            _src_sub[_col_estado].str.contains(r[0], na=False).any() for r in rangos
        )
        if tiene_datos:
            st.markdown("**Distribución por Sub-categoría de Diagnóstico — Subexplotados**")
            dcols = st.columns(3)
            for col_w, (rango, color) in zip(dcols, rangos):
                sub_df = _src_sub[
                    _src_sub[_col_estado].str.contains(rango, na=False, regex=False)
                    & _src_sub["DIAG_SUBCATEGORIA"].notna()
                    & (_src_sub["DIAG_SUBCATEGORIA"].astype(str).str.strip() != "")
                ].copy()
                with col_w:
                    label_corto = (
                        rango.replace("Subexplotado bajo entre 201-400m",    "Subexplot. Bajo (201-400m)")
                             .replace("Subexplotado moderado entre 401-600m", "Subexplot. Moderado (401-600m)")
                             .replace("Subexplotado alto > 600m",             "Subexplot. Alto (>600m)")
                    )
                    st.markdown(
                        f"<div style='font-size:0.78rem; font-weight:700; color:{color}; margin-bottom:4px'>"
                        f"{label_corto}</div>",
                        unsafe_allow_html=True
                    )
                    if sub_df.empty:
                        st.info("Sin diagnóstico registrado.")
                    else:
                        sub_cnt = (
                            sub_df.groupby("DIAG_SUBCATEGORIA", as_index=False)
                            .size()
                            .rename(columns={"size": "Total"})
                            .sort_values("Total", ascending=True)
                        )
                        fig_sub = px.bar(
                            sub_cnt, x="Total", y="DIAG_SUBCATEGORIA",
                            orientation="h", text="Total",
                            color_discrete_sequence=[color],
                        )
                        fig_sub.update_traces(textposition="outside", textfont_size=11, cliponaxis=False)
                        fig_sub.update_layout(
                            plot_bgcolor="white", paper_bgcolor="white",
                            height=max(220, len(sub_cnt) * 55 + 40),
                            margin=dict(l=210, r=50, t=10, b=30),
                            xaxis_title="Pozos", yaxis_title="",
                            font=dict(family="Inter", size=11),
                            showlegend=False,
                            yaxis=dict(tickfont=dict(size=10), ticklen=0),
                        )
                        st.plotly_chart(fig_sub, use_container_width=True)
                        # Listado de pozos por subcategoría
                        with st.expander(f"Ver pozos ({len(sub_df)})"):
                            cols_exp = [c for c in [
                                "NOMBRE_POZO", "BATERIA", "METODO",
                                "ULT_SUMERGE", "DIAG_SUBCATEGORIA"
                            ] if c in sub_df.columns]
                            # Si no hay ULT_SUMERGE, intentar SUMERGENCIA desde maestro
                            if "ULT_SUMERGE" not in cols_exp and "NOMBRE_POZO" in sub_df.columns:
                                sub_df = sub_df.merge(
                                    df_tab1[["NOMBRE_POZO", "SUMERGENCIA"]].rename(
                                        columns={"SUMERGENCIA": "SUMERG"}
                                    ),
                                    on="NOMBRE_POZO", how="left"
                                )
                                cols_exp = [c for c in cols_exp + ["SUMERG"] if c in sub_df.columns]
                            tabla_pozos = (
                                sub_df[cols_exp]
                                .sort_values("DIAG_SUBCATEGORIA")
                                .reset_index(drop=True)
                            )
                            tabla_pozos.index += 1
                            st.dataframe(tabla_pozos, use_container_width=True, height=220)

    # ══ COBERTURA DE ACCIONES — SUB-EXPLOTADOS ══════════════════
    # Cruza Historico_Diagnosticos con Acciones:
    # cuenta como "con acción" si existe ≥1 acción POSTERIOR a la fecha del diagnóstico.
    # ════════════════════════════════════════════════════════════

    if _usa_diag and "DIAG_SUBCATEGORIA" in ult_diag_t1.columns:
        st.markdown("---")
        st.markdown("**Cobertura de Acciones — Sub-explotados**")
        st.caption(
            "Se considera que un pozo tiene acción si existe al menos una acción registrada "
            "con fecha **posterior** a la del diagnóstico activo."
        )

        _acc = acciones.copy()

        def _norm_col(c):
            """Normaliza nombre de columna para comparación: mayúsculas sin tildes."""
            return (str(c).upper()
                    .replace("Á","A").replace("É","E").replace("Í","I")
                    .replace("Ó","O").replace("Ú","U").replace("Ü","U"))

        # ── Detectar columna de nombre de pozo en Acciones ──
        _col_pa = next((c for c in _acc.columns if any(
            p in _norm_col(c) for p in ["NOMBRE_CORTO", "NOMBRE_POZO", "NOMBRE", "POZO"]
        )), None)
        # ── Detectar columna de fecha en Acciones (incluye "Fecha solicitado") ──
        _col_fa = next((c for c in _acc.columns if any(
            p in _norm_col(c) for p in ["FECHA", "DATE", "DIA", "INGRESO", "SOLICIT"]
        )), None)

        _sin_fecha = _col_fa is None   # Acciones no tiene columna de fecha → conteo simple

        if _acc.empty or not _col_pa:
            st.info(
                f"No se pudo construir el cruce (columna de pozo no encontrada). "
                f"Todas las columnas en Acciones: `{list(_acc.columns)}`"
            )
        else:
            # Normalizar Acciones — columna pozo siempre; fecha si existe
            _cols_usar = [_col_pa] + ([_col_fa] if _col_fa else [])
            _acc = _acc[_cols_usar].copy()
            _acc.columns = ["NOMBRE_POZO"] + (["FECHA_ACCION"] if _col_fa else [])
            _acc["NOMBRE_POZO"] = _acc["NOMBRE_POZO"].astype(str).str.strip().str.upper()
            _acc = _acc[_acc["NOMBRE_POZO"].str.replace(" ","") != "NAN"]

            if _col_fa:
                _acc["FECHA_ACCION"] = pd.to_datetime(_acc["FECHA_ACCION"], errors="coerce")
                _acc = _acc.dropna(subset=["FECHA_ACCION"])
                # Última acción por pozo
                _ult_acc = (
                    _acc.groupby("NOMBRE_POZO")["FECHA_ACCION"].max()
                    .reset_index()
                    .rename(columns={"FECHA_ACCION": "FECHA_ULT_ACCION"})
                )
            else:
                # Sin columna de fecha: cualquier acción registrada cuenta
                _ult_acc = (
                    _acc[["NOMBRE_POZO"]].drop_duplicates()
                    .assign(FECHA_ULT_ACCION=pd.NaT)
                )
                st.caption(
                    "⚠️ La hoja Acciones no tiene columna de fecha. "
                    f"Columnas disponibles: {list(acciones.columns)}. "
                    "Se cuenta si el pozo tiene CUALQUIER acción registrada (sin filtro de fecha)."
                )

            # Sub-explotados desde el diagnóstico (todos, sin importar si tienen sub-categoría)
            _sub = ult_diag_t1[
                ult_diag_t1["DIAG_ESTADO_OP"].str.contains("Subexplotado", na=False)
            ].copy()

            # Normalizar NOMBRE_POZO a uppercase para que el merge con Acciones funcione
            _sub["NOMBRE_POZO"] = _sub["NOMBRE_POZO"].astype(str).str.strip().str.upper()
            if "DIAG_FECHA" in _sub.columns:
                _sub["DIAG_FECHA"] = pd.to_datetime(_sub["DIAG_FECHA"], errors="coerce")

            _sub = _sub.merge(_ult_acc, on="NOMBRE_POZO", how="left")

            # Flag: acción POSTERIOR al diagnóstico (si hay fecha en Acciones) o existente (sin fecha)
            if not _sin_fecha and "DIAG_FECHA" in _sub.columns:
                _sub["TIENE_ACCION"] = (
                    _sub["FECHA_ULT_ACCION"].notna()
                    & (_sub["FECHA_ULT_ACCION"] > _sub["DIAG_FECHA"])
                )
            else:
                # Sin fecha en Acciones: cuenta si el pozo tiene cualquier acción
                _sub["TIENE_ACCION"] = _sub["FECHA_ULT_ACCION"].notna()

            # ── Construir tabla resumen ──
            # Usamos keywords únicos (bajo/moderado/alto) para capturar variaciones de texto
            _rangos_acc = [
                ("bajo",     "201–400 m",   C_PRIM),
                ("moderado", "401–600 m",   "#E67E22"),
                ("alto",     "> 600 m",     C_RED),
            ]

            # Métrica global
            _n_sub_total  = len(_sub)
            _n_con_accion = int(_sub["TIENE_ACCION"].sum())
            _n_sin_accion = _n_sub_total - _n_con_accion
            _pct_global   = f"{100*_n_con_accion/_n_sub_total:.0f}%" if _n_sub_total else "–"

            _ma, _mb, _mc = st.columns(3)
            for _col_m, _val, _lbl, _col_c in [
                (_ma, _n_sub_total,  "Sub-explotados",  C_NAVY),
                (_mb, _n_con_accion, "Con Acción",      C_PRIM),
                (_mc, _n_sin_accion, "Sin Acción",      C_RED),
            ]:
                with _col_m:
                    st.markdown(
                        f"<div style='background:{C_BG};border-radius:8px;padding:12px 16px;"
                        f"text-align:center;margin-bottom:10px'>"
                        f"<div style='font-size:1.6rem;font-weight:800;color:{_col_c}'>{_val}</div>"
                        f"<div style='font-size:0.75rem;color:#555;margin-top:2px'>{_lbl}</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

            # ── Tablas por rango ──
            _rcols = st.columns(3)
            for _rc, (_rkey, _rlabel, _rcolor) in zip(_rcols, _rangos_acc):
                _df_r = _sub[
                    _sub["DIAG_ESTADO_OP"].str.contains(_rkey, na=False, case=False)
                ].copy()

                with _rc:
                    _n_r = len(_df_r)
                    _c_r = int(_df_r["TIENE_ACCION"].sum())
                    _pct_r = f"{100*_c_r/_n_r:.0f}%" if _n_r else "–"
                    st.markdown(
                        f"<div style='font-weight:700;color:{_rcolor};"
                        f"font-size:0.85rem;margin-bottom:6px'>"
                        f"{_rlabel} — {_n_r} pozos | {_c_r} con acción ({_pct_r})</div>",
                        unsafe_allow_html=True
                    )

                    if _df_r.empty:
                        st.info("Sin pozos en este rango.")
                    else:
                        # Agrupar por sub-categoría (incluye "Sin categoría" para NaN)
                        _df_r["DIAG_SUBCATEGORIA"] = _df_r["DIAG_SUBCATEGORIA"].fillna("Sin categoría").replace("", "Sin categoría")
                        _filas_r = []
                        for _sc, _grp in _df_r.groupby("DIAG_SUBCATEGORIA", dropna=False):
                            _t = len(_grp)
                            _c = int(_grp["TIENE_ACCION"].sum())
                            _s = _t - _c
                            _p = f"{100*_c/_t:.0f}%" if _t else "–"
                            _filas_r.append({
                                "Sub-categoría": _sc,
                                "Total": _t,
                                "✅ Con Acción": _c,
                                "❌ Sin Acción": _s,
                                "Cobertura": _p,
                            })
                        _tbl_r = pd.DataFrame(_filas_r).sort_values("Total", ascending=False)
                        _tbl_r.index = range(1, len(_tbl_r) + 1)
                        st.dataframe(
                            _tbl_r,
                            use_container_width=True,
                            height=min(400, len(_tbl_r) * 42 + 50),
                        )

                        # Expander: listado de pozos sin acción
                        _sin = _df_r[~_df_r["TIENE_ACCION"]].copy()
                        if not _sin.empty:
                            with st.expander(f"Sin acción ({len(_sin)} pozos)"):
                                _cols_sin = [c for c in [
                                    "NOMBRE_POZO", "BATERIA", "DIAG_SUBCATEGORIA",
                                    "ULT_SUMERGE", "DIAG_FECHA"
                                ] if c in _sin.columns]
                                st.dataframe(
                                    _sin[_cols_sin].sort_values("DIAG_SUBCATEGORIA").reset_index(drop=True),
                                    use_container_width=True, height=200
                                )

    st.markdown("<div style='margin-top:18px'></div>", unsafe_allow_html=True)

    # ── Antigüedad de mediciones por batería ──
    st.markdown("**Pozos con más de 30 días sin medición — por Batería**")
    _UMBRAL_DIAS = 30
    _hoy_med = pd.Timestamp.today().normalize()
    _cols_med = {
        "FECHA_ULT_CONTROL": ("Sin Control", C_RED),
        "FECHA_ULT_DINA":    ("Sin Dina",    C_NAVY),
        "FECHA_ULT_NIVEL":   ("Sin Nivel",   C_PRIM),
    }
    _series_med = []
    _df_med_base = df_tab1.copy() if "BATERIA_FINAL" in df_tab1.columns else pd.DataFrame()
    for _col_f, (_label, _col_color) in _cols_med.items():
        if not _df_med_base.empty and _col_f in _df_med_base.columns:
            _tmp = _df_med_base[["BATERIA_FINAL", _col_f]].copy()
            _tmp[_col_f] = pd.to_datetime(_tmp[_col_f], errors="coerce")
            _tmp["_dias"] = (_hoy_med - _tmp[_col_f]).dt.days
            _cnt = (
                _tmp[_tmp["_dias"] > _UMBRAL_DIAS]
                .groupby("BATERIA_FINAL", as_index=False)
                .size()
                .rename(columns={"size": "Total", "BATERIA_FINAL": "Batería"})
            )
            _cnt["Tipo"] = _label
            _series_med.append((_cnt, _col_color))

    if _series_med:
        _df_plot_med = pd.concat([s for s, _ in _series_med], ignore_index=True)
        _df_plot_med["Batería"] = _df_plot_med["Batería"].astype(str)
        _bats_ord = sorted(
            _df_plot_med["Batería"].unique(),
            key=lambda x: int(x) if x.isdigit() else 9999,
        )
        _cmap_med = {_label: _col_color for _, (_label, _col_color) in _cols_med.items()}
        fig_med = px.bar(
            _df_plot_med,
            x="Batería", y="Total",
            color="Tipo",
            barmode="group",
            text="Total",
            color_discrete_map=_cmap_med,
            category_orders={"Batería": _bats_ord},
        )
        fig_med.update_traces(textposition="outside", textfont_size=10, cliponaxis=False)
        fig_med.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            height=320,
            margin=dict(l=10, r=10, t=20, b=40),
            xaxis_title="Batería",
            yaxis_title=f"Pozos > {_UMBRAL_DIAS} días sin medición",
            font=dict(family="Inter", size=11),
            legend=dict(orientation="h", y=1.08, title_text=""),
        )
        st.plotly_chart(fig_med, use_container_width=True)
    else:
        st.info("No hay datos de fechas de medición disponibles.")

    st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)

    # ── Gráfico Riesgo Operacional ──
    if "RIESGO_OPERACIONAL" in df_tab1.columns and len(df_tab1) > 0:
        r1, r2 = st.columns(2)
        with r1:
            st.markdown("**Distribución por Riesgo Operacional**")
            riesgo_cnt = (
                df_tab1.groupby("RIESGO_OPERACIONAL", as_index=False)
                .size()
                .rename(columns={"size": "Total"})
            )
            color_map = {
                "Critico": C_RED,
                "Alto":    "#E67E22",
                "Medio":   C_NAVY,
                "Bajo":    C_PRIM,
            }
            fig_riesgo = px.bar(
                riesgo_cnt, x="RIESGO_OPERACIONAL", y="Total",
                text="Total",
                color="RIESGO_OPERACIONAL",
                color_discrete_map=color_map,
            )
            fig_riesgo.update_traces(textposition="outside", textfont_size=13, cliponaxis=False)
            fig_riesgo.update_layout(
                plot_bgcolor="white", paper_bgcolor="white",
                height=280, margin=dict(l=10, r=10, t=50, b=30),
                xaxis_title="Riesgo Operacional", yaxis_title="Total Pozos",
                font=dict(family="Inter", size=12),
                showlegend=False,
            )
            st.plotly_chart(fig_riesgo, use_container_width=True)

        with r2:
            st.markdown("**Distribución por Sistema Extractivo**")
            if "SISTEMA_EXTRACTIVO" in df_tab1.columns:
                sis_cnt = (
                    df_tab1.groupby("SISTEMA_EXTRACTIVO", as_index=False)
                    .size()
                    .rename(columns={"size": "Total"})
                    .sort_values("Total", ascending=False)
                )
                fig_sis = px.bar(
                    sis_cnt, x="SISTEMA_EXTRACTIVO", y="Total",
                    text="Total",
                    color_discrete_sequence=[C_PRIM],
                )
                fig_sis.update_traces(textposition="outside", textfont_size=13, cliponaxis=False)
                fig_sis.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white",
                    height=280, margin=dict(l=10, r=10, t=50, b=30),
                    xaxis_title="Sistema Extractivo", yaxis_title="Total Pozos",
                    font=dict(family="Inter", size=12),
                    showlegend=False,
                )
                st.plotly_chart(fig_sis, use_container_width=True)

    # ── Tabla filtrable ──
    st.markdown("---")
    st.markdown("**Listado de Pozos**")

    COLS_TABLA = [
        "NOMBRE_POZO", "ZONA", "BATERIA_FINAL", "SISTEMA_EXTRACTIVO",
        "ESTADO_OPERATIVO", "RIESGO_OPERACIONAL",
        "SUMERGENCIA", "BBA_LLENADO_DE_BOMBA",
        "BBA_EFICIENCIA_VOLUMETRICA", "AIBRE_SOLICITACION_DE_ESTRUCT",
        "EVENTOS_12M", "RIESGO_HISTORICO",
    ]
    RENAME_TABLA = {
        "BATERIA_FINAL":               "BATERÍA",
        "SISTEMA_EXTRACTIVO":          "SISTEMA",
        "ESTADO_OPERATIVO":            "ESTADO OPERATIVO",
        "RIESGO_OPERACIONAL":          "RIESGO",
        "SUMERGENCIA":                 "SUMERGENCIA (m)",
        "BBA_LLENADO_DE_BOMBA":        "LLENADO BBA (%)",
        "BBA_EFICIENCIA_VOLUMETRICA":  "EFIC VOL (%)",
        "AIBRE_SOLICITACION_DE_ESTRUCT": "SOL EST (%)",
    }

    cols_show = [c for c in COLS_TABLA if c in df_tab1.columns]
    df_tabla  = df_tab1[cols_show].rename(columns=RENAME_TABLA).reset_index(drop=True)

    # Redondear numéricos
    for c in ["SUMERGENCIA (m)", "LLENADO BBA (%)", "EFIC VOL (%)", "SOL EST (%)"]:
        if c in df_tabla.columns:
            df_tabla[c] = pd.to_numeric(df_tabla[c], errors="coerce").round(1)

    styled = (
        df_tabla.style
        .map(color_riesgo, subset=["RIESGO"])
        if "RIESGO" in df_tabla.columns
        else df_tabla.style
    )
    st.dataframe(styled, use_container_width=True, height=380, hide_index=True)


# ══════════════════════════════════════════════════════════════
#  TAB 2  —  DETALLE DE POZO
# ══════════════════════════════════════════════════════════════

@st.fragment
def tab2_detalle(idx, maestro, df, last_update):
    """Fragment: solo se re-ejecuta cuando cambia el pozo o un widget interno."""

    # ── Header ──
    st.markdown(f"""
    <div class="qe-header">
      <div>
        <h1>Análisis Detallado de Pozo</h1>
        <div class="subtitle">Última actualización datos: {last_update}</div>
      </div>
      <div class="brand">Quintana Energy</div>
    </div>
    """, unsafe_allow_html=True)

    # ── Selector de pozo ──
    pozos_lista = sorted(df["NOMBRE_POZO"].dropna().unique().tolist()) if "NOMBRE_POZO" in df.columns else []

    if not pozos_lista:
        st.warning("No hay pozos disponibles con los filtros seleccionados.")
        st.stop()

    col_pozo, _ = st.columns([2, 3])
    with col_pozo:
        pozo_sel = st.selectbox("🔍 Seleccionar Pozo", options=pozos_lista, key="pozo_sel")

    # Normalizar nombre para lookups en idx (idx keys son siempre MAYÚSCULAS)
    _psel_up = pozo_sel.strip().upper()

    # Datos del pozo seleccionado
    mask = maestro["NOMBRE_POZO"] == pozo_sel
    if not mask.any():
        st.warning("No se encontraron datos para el pozo seleccionado.")
        st.stop()
    row = maestro[mask].iloc[0]

    st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)

    # ══ FILA 1: Batería | Sistema | Estado | Alerta ══
    ic1, ic2, ic3, ic4 = st.columns([1, 1.5, 1.2, 2.2])

    with ic1:
        st.markdown(f"""
        <div class="info-card">
          <div class="info-card-label">Batería</div>
          <div class="info-card-value">{safe_val(row.get("BATERIA_FINAL"), decimals=0)}</div>
        </div>""", unsafe_allow_html=True)

    with ic2:
        metodo_disp = str(row.get("METODO_EXTRACCION") or "–")
        st.markdown(f"""
        <div class="info-card">
          <div class="info-card-label">Sistema Extractivo</div>
          <div class="info-card-value">{metodo_disp}</div>
        </div>""", unsafe_allow_html=True)

    with ic3:
        riesgo_op = str(row.get("RIESGO_OPERACIONAL") or "–")
        color_riesgo_card = {
            "Critico": C_RED,
            "Alto":    "#C47A00",
            "Medio":   C_NAVY,
            "Bajo":    C_PRIM,
        }.get(riesgo_op, C_TEXT)
        st.markdown(f"""
        <div class="info-card">
          <div class="info-card-label">Estado Operativo</div>
          <div class="info-card-value" style="color:{color_riesgo_card}">{riesgo_op}</div>
        </div>""", unsafe_allow_html=True)

    with ic4:
        alerta_txt = str(row.get("ALERTA") or "Sin alerta relevante")
        hay_alerta = alerta_txt != "Sin alerta relevante" and alerta_txt != "–"
        if hay_alerta:
            alertas_html = alerta_txt.replace(" | ", "<br>")
            st.markdown(f"""
            <div class="alerta-card">
              <div class="alerta-card-label">⚠ Alerta Operativa</div>
              <div class="alerta-card-value">{alertas_html}</div>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="alerta-ok">
              <div class="alerta-ok-label">✓ Sin Alertas</div>
              <div class="alerta-ok-value">Condición operativa normal</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:14px'></div>", unsafe_allow_html=True)

    # ══ FILA 2: Producción + Tendencias | Info Adicional ══
    col_main, col_info = st.columns([3.2, 1])

    with col_main:
        # KPIs de producción
        prod_bruta = pd.to_numeric(row.get("TOTAL_LIQUIDO"), errors="coerce")
        prod_neta  = pd.to_numeric(row.get("PROD_OIL"), errors="coerce")
        pct_agua   = (
            (prod_bruta - prod_neta) / prod_bruta * 100
            if (pd.notna(prod_bruta) and pd.notna(prod_neta) and prod_bruta > 0)
            else np.nan
        )

        pk1, pk2, pk3 = st.columns(3)
        for col_w, lbl, val, sfx in [
            (pk1, "Producción Bruta",  prod_bruta, " m³/d"),
            (pk2, "Producción Neta",   prod_neta,  " m³/d"),
            (pk3, "% Agua",            pct_agua,   " %"),
        ]:
            with col_w:
                st.markdown(f"""
                <div class="prod-card">
                  <div class="prod-card-label">{lbl}</div>
                  <div class="prod-card-value">{safe_val(val)}{sfx}</div>
                </div>""", unsafe_allow_html=True)

        st.markdown("<div style='margin-top:12px'></div>", unsafe_allow_html=True)

        # Tendencias de producción — ya procesadas en el cache
        prod_pozo = idx["prod"].get(_psel_up, pd.DataFrame())

        tc1, tc2 = st.columns(2)

        with tc1:
            st.markdown("**Tendencia Producción**")
            if not prod_pozo.empty:
                fig_prod = go.Figure()
                if "TOTAL_LIQUIDO" in prod_pozo.columns:
                    fig_prod.add_trace(go.Scatter(
                        x=prod_pozo["FECHA"], y=prod_pozo["TOTAL_LIQUIDO"],
                        name="Total Líquido",
                        mode="lines+markers",
                        line=dict(color=C_PRIM, width=2),
                        marker=dict(size=4, color=C_PRIM),
                        yaxis="y1",
                    ))
                if "PROD_OIL" in prod_pozo.columns:
                    fig_prod.add_trace(go.Scatter(
                        x=prod_pozo["FECHA"], y=prod_pozo["PROD_OIL"],
                        name="Prod Oil",
                        mode="lines+markers",
                        line=dict(color="#27AE60", width=2, dash="dot"),
                        marker=dict(size=4, color="#27AE60"),
                        yaxis="y2",
                    ))
                fig_prod.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white",
                    height=240, margin=dict(l=10, r=40, t=10, b=10),
                    legend=dict(orientation="h", y=1.15, font_size=10),
                    font=dict(family="Inter", size=11),
                    xaxis_title="FECHA",
                    yaxis=dict(title="Total Líquido", title_font=dict(color=C_PRIM)),
                    yaxis2=dict(
                        title="Prod Oil", title_font=dict(color="#27AE60"),
                        overlaying="y", side="right",
                    ),
                )
                st.plotly_chart(fig_prod, use_container_width=True)
            else:
                st.info("Sin datos de producción.")

        with tc2:
            st.markdown("**Tendencia Corte de Agua**")
            if not prod_pozo.empty and "TOTAL_LIQUIDO" in prod_pozo.columns and "PROD_OIL" in prod_pozo.columns:
                prod_pozo["PCT_AGUA"] = (
                    (prod_pozo["TOTAL_LIQUIDO"] - prod_pozo["PROD_OIL"])
                    / prod_pozo["TOTAL_LIQUIDO"] * 100
                ).clip(0, 100)
                fig_agua = px.line(
                    prod_pozo.dropna(subset=["PCT_AGUA"]),
                    x="FECHA", y="PCT_AGUA",
                    color_discrete_sequence=[C_NAVY],
                    markers=True,
                )
                fig_agua.update_traces(marker_size=4)
                fig_agua.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white",
                    height=240, margin=dict(l=10, r=10, t=10, b=10),
                    yaxis_title="% Agua", yaxis_ticksuffix="%",
                    font=dict(family="Inter", size=11), showlegend=False,
                )
                st.plotly_chart(fig_agua, use_container_width=True)
            else:
                st.info("Sin datos de producción suficientes.")

    with col_info:
        # Información técnica del pozo
        st.markdown("**Información Adicional**")

        tipo_bomba_aib = str(row.get("AIB_MARCA_Y_DESC_API") or "–")
        tipo_bomba_fab = str(row.get("FABRICANTE_BOMBA") or "–")
        modelo_bomba   = str(row.get("MODELO_BOMBA") or "–")

        def _dias_desde(fecha_val):
            """Devuelve string con días transcurridos desde la fecha, o '–'."""
            try:
                f = pd.to_datetime(fecha_val, errors="coerce")
                if pd.isna(f):
                    return "–"
                dias = (pd.Timestamp.today().normalize() - f.normalize()).days
                return f"{dias} días"
            except Exception:
                return "–"

        info_rows = {
            "Carrera":          safe_val(row.get("AIB_CARRERA_DEL_VASTAGO"), decimals=0, suffix=" in"),
            "Tipo AIB":         tipo_bomba_aib[:22] + "…" if len(tipo_bomba_aib) > 22 else tipo_bomba_aib,
            "Fabricante":       tipo_bomba_fab,
            "Modelo Bomba":     modelo_bomba[:20] + "…" if len(modelo_bomba) > 20 else modelo_bomba,
            "Polea Motor":      safe_val(row.get("MOTOR_DIAMETRO_POLEA"), decimals=0),
            "Prof. Bomba":      safe_val(row.get("PROF_ENT_BOMBA"), decimals=0, suffix=" m"),
            "Régimen":          safe_val(row.get("REGIMEN_PRIM"), suffix=" GPM"),
            "RPM Motor":        safe_val(row.get("MOTOR_RPM"), decimals=0),
            "Último Control":   safe_date(row.get("FECHA_ULT_CONTROL")),
            "Días sin control": _dias_desde(row.get("FECHA_ULT_CONTROL")),
            "Último Dina":      safe_date(row.get("FECHA_ULT_DINA")),
            "Último Nivel":     safe_date(row.get("FECHA_ULT_NIVEL")),
            "Días sin nivel":   _dias_desde(row.get("FECHA_ULT_NIVEL")),
        }
        df_info = pd.DataFrame(info_rows.items(), columns=["Atributo", "Valor"])
        st.dataframe(df_info, use_container_width=True, hide_index=True, height=460)

    st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)

    # ══ FILA 3: KPIs operativos ══
    op_cols = st.columns(5)
    op_data = [
        ("Sumergencia",     row.get("SUMERGENCIA"),                  " m"),
        ("Llenado Bba",     row.get("BBA_LLENADO_DE_BOMBA"),         " %"),
        ("Eficiencia Vol",  row.get("BBA_EFICIENCIA_VOLUMETRICA"),   " %"),
        ("Sol. Estructural",row.get("AIBRE_SOLICITACION_DE_ESTRUCT")," %"),
        ("Torque Red",      row.get("AIBRR_PORCENTAJE"),             " %"),
    ]
    for col_w, (lbl, val, sfx) in zip(op_cols, op_data):
        with col_w:
            num_val = pd.to_numeric(val, errors="coerce")
            display = safe_val(num_val, decimals=0) + sfx if pd.notna(num_val) else "–"
            st.markdown(f"""
            <div class="kpi-sm">
              <div class="kpi-sm-label">{lbl}</div>
              <div class="kpi-sm-value">{display}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:14px'></div>", unsafe_allow_html=True)

    # ══ FILA 4: Tendencia Sumergencia | Tendencia Dinamometría ══
    # Datos — ya procesados en el cache
    niv_pozo  = idx["niv"].get(_psel_up,  pd.DataFrame())
    dina_pozo = idx["dina"].get(_psel_up, pd.DataFrame())

    td1, td2 = st.columns(2)

    with td1:
        st.markdown("**Tendencia Sumergencia**")
        if not niv_pozo.empty:
            fig_sum = px.line(
                niv_pozo, x="FECHA_NVL", y="SUMERGENCIA",
                color_discrete_sequence=[C_NAVY],
                markers=True,
            )
            fig_sum.update_traces(marker_size=4)
            fig_sum.update_layout(
                plot_bgcolor="white", paper_bgcolor="white",
                height=250, margin=dict(l=10, r=10, t=10, b=10),
                xaxis_title="Fecha", yaxis_title="Sumergencia (m)",
                font=dict(family="Inter", size=11), showlegend=False,
            )
            st.plotly_chart(fig_sum, use_container_width=True)
        else:
            st.info("Sin datos de sumergencia.")

    with td2:
        st.markdown("**Tendencia Dinamometría**")
        if not dina_pozo.empty:
            fig_dyn = go.Figure()
            if "BBA_LLENADO_DE_BOMBA" in dina_pozo.columns:
                fig_dyn.add_trace(go.Scatter(
                    x=dina_pozo["FECHA_DINA"], y=dina_pozo["BBA_LLENADO_DE_BOMBA"],
                    name="Llenado BBA (%)",
                    mode="lines+markers",
                    line=dict(color=C_NAVY, width=2),
                    marker=dict(size=4, color=C_NAVY),
                ))
            if "BBA_EFICIENCIA_VOLUMETRICA" in dina_pozo.columns:
                fig_dyn.add_trace(go.Scatter(
                    x=dina_pozo["FECHA_DINA"], y=dina_pozo["BBA_EFICIENCIA_VOLUMETRICA"],
                    name="Eficiencia Vol (%)",
                    mode="lines+markers",
                    line=dict(color=C_RED, width=2),
                    marker=dict(size=4, color=C_RED),
                ))
            fig_dyn.update_layout(
                plot_bgcolor="white", paper_bgcolor="white",
                height=250, margin=dict(l=10, r=10, t=10, b=10),
                legend=dict(orientation="h", y=1.18, font_size=10),
                xaxis_title="Fecha", yaxis_title="(%)",
                font=dict(family="Inter", size=11),
            )
            st.plotly_chart(fig_dyn, use_container_width=True)
        else:
            st.info("Sin datos de dinamometría.")

    # ══ CARTAS DINAMOMÉTRICAS ══
    st.markdown('<div class="section-title">Últimas 3 Cartas Dinamométricas</div>', unsafe_allow_html=True)

    cartas_pozo = idx["cartas"].get(_psel_up, pd.DataFrame())

    def mostrar_cartas(df_tipo, tipo_label, palette):
        if df_tipo.empty or "N_DINA" not in df_tipo.columns:
            st.info(f"Sin cartas de {tipo_label}.")
            return
        n_dinas = sorted(df_tipo["N_DINA"].dropna().unique())[:3]
        if not n_dinas:
            st.info(f"Sin cartas de {tipo_label}.")
            return
        cols_w = st.columns(len(n_dinas))
        for i, n in enumerate(n_dinas):
            carta_df = df_tipo[df_tipo["N_DINA"] == n].sort_values("PUNTO")
            fecha_c  = pd.to_datetime(carta_df["FECHA_DINA"].iloc[0], errors="coerce") if "FECHA_DINA" in carta_df.columns else None
            titulo   = "Último dina" if i == 0 else f"Dina -{i}"
            fecha_s  = fecha_c.strftime("%d/%m/%Y %H:%M") if fecha_c and not pd.isna(fecha_c) else ""
            with cols_w[i]:
                fig_c = go.Figure()
                fig_c.add_trace(go.Scatter(
                    x=carta_df["X"], y=carta_df["Y"],
                    mode="lines",
                    line=dict(color=palette[i], width=2.5),
                    showlegend=False,
                ))
                fig_c.update_layout(
                    title=dict(
                        text=f"<b>{titulo}</b><br><sup>{fecha_s}</sup>",
                        font_size=12, x=0.5, xanchor="center",
                    ),
                    plot_bgcolor="white", paper_bgcolor="white",
                    height=230, margin=dict(l=10, r=10, t=55, b=30),
                    xaxis_title="Posición / carrera",
                    yaxis_title="Carga",
                    font=dict(family="Inter", size=10),
                )
                st.plotly_chart(fig_c, use_container_width=True)

    if cartas_pozo.empty:
        st.info("Sin cartas dinamométricas disponibles para este pozo.")
    else:
        tiene_tipo = "TIPO_CARTA" in cartas_pozo.columns

        # ── Cartas de Superficie (CS) — arriba ──
        if tiene_tipo:
            cs = cartas_pozo[cartas_pozo["TIPO_CARTA"] == "CS"]
            if not cs.empty:
                st.markdown("**📈 Carta de Superficie (CS)**")
                mostrar_cartas(cs, "Superficie", [C_NAVY, "#8E44AD", "#27AE60"])

        # ── Cartas de Fondo (CF) — abajo ──
        st.markdown("**📉 Carta de Fondo (CF)**")
        cf = cartas_pozo[cartas_pozo["TIPO_CARTA"] == "CF"] if tiene_tipo else cartas_pozo
        mostrar_cartas(cf, "Fondo", [C_PRIM, "#E67E22", C_NAVY])

    # ── Diagnóstico de Cartas ──
    diag_pozo = idx["diag"].get(_psel_up, pd.DataFrame())
    if not diag_pozo.empty:
        cols_d = [c for c in ["N_DINA", "FECHA_DINA", "DIAGNOSTICO_CARTA",
                               "COMENTARIO_CARTA", "CONFIANZA_CARTA"] if c in diag_pozo.columns]
        rename_d = {
            "DIAGNOSTICO_CARTA": "Diagnóstico",
            "COMENTARIO_CARTA":  "Comentario",
            "CONFIANZA_CARTA":   "Confianza",
            "FECHA_DINA":        "Fecha",
        }
        st.dataframe(
            diag_pozo[cols_d].rename(columns=rename_d).reset_index(drop=True),
            use_container_width=True, hide_index=True,
        )

    # ══ HISTORIAL DE DIAGNÓSTICOS OPERATIVOS ══
    st.markdown('<div class="section-title">Historial de Diagnósticos Operativos</div>', unsafe_allow_html=True)

    diag_op_pozo = idx["diag_op"].get(_psel_up, pd.DataFrame())

    # Fallback: si la hoja Historico_Diagnosticos aún no fue generada,
    # usar las columnas DIAG_ que ya están en el maestro (último diagnóstico)
    if diag_op_pozo.empty:
        diag_txt_fb = str(row.get("DIAG_TEXTO") or "").strip()
        if diag_txt_fb and diag_txt_fb not in ("nan", "–"):
            diag_op_pozo = pd.DataFrame([{
                "DIAG_FECHA":        row.get("DIAG_FECHA"),
                "DIAG_ESTADO_OP":    str(row.get("DIAG_ESTADO_OP")    or ""),
                "DIAG_SUBCATEGORIA": str(row.get("DIAG_SUBCATEGORIA") or ""),
                "DIAG_TEXTO":        diag_txt_fb,
                "DIAG_OBSERVACION":  str(row.get("DIAG_OBSERVACION")  or ""),
            }])

    if not diag_op_pozo.empty:
        if "DIAG_FECHA" in diag_op_pozo.columns:
            diag_op_pozo = diag_op_pozo.copy()
            diag_op_pozo["DIAG_FECHA"] = pd.to_datetime(diag_op_pozo["DIAG_FECHA"], errors="coerce")
            diag_op_pozo = diag_op_pozo.sort_values("DIAG_FECHA", ascending=False)

        cols_diag_op = [c for c in [
            "DIAG_FECHA", "DIAG_ESTADO_OP", "DIAG_SUBCATEGORIA",
            "DIAG_TEXTO", "DIAG_OBSERVACION"
        ] if c in diag_op_pozo.columns]

        rename_diag_op = {
            "DIAG_FECHA":        "Fecha",
            "DIAG_ESTADO_OP":    "Estado",
            "DIAG_SUBCATEGORIA": "Sub-categoría",
            "DIAG_TEXTO":        "Diagnóstico",
            "DIAG_OBSERVACION":  "Observación / Recomendación",
        }
        st.dataframe(
            diag_op_pozo[cols_diag_op].rename(columns=rename_diag_op).reset_index(drop=True),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("Sin diagnósticos registrados para este pozo.")

    st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)

    # ══ ÚLTIMAS ACCIONES ══
    st.markdown('<div class="section-title">Últimas Acciones</div>', unsafe_allow_html=True)

    acc_pozo = pd.DataFrame()
    if not acciones.empty:
        acc_norm = acciones.copy()
        acc_norm.columns = [str(c).strip() for c in acc_norm.columns]
        pozo_col_acc = next(
            (c for c in acc_norm.columns
             if c.upper() in ["NOMBRE_POZO", "POZO", "WELL", "NOMBRE"]),
            None
        )
        if pozo_col_acc:
            acc_pozo = acc_norm[
                acc_norm[pozo_col_acc].astype(str).str.strip().str.upper() == pozo_sel.upper()
            ].copy()
            if not acc_pozo.empty:
                fecha_col_acc = next(
                    (c for c in acc_pozo.columns if "FECHA" in c.upper()), None
                )
                if fecha_col_acc:
                    acc_pozo[fecha_col_acc] = pd.to_datetime(acc_pozo[fecha_col_acc], errors="coerce")
                    acc_pozo = acc_pozo.sort_values(fecha_col_acc, ascending=False)
                show_cols_acc = [c for c in acc_pozo.columns if "Unnamed" not in str(c)]
                st.dataframe(
                    acc_pozo[show_cols_acc].head(5).reset_index(drop=True),
                    use_container_width=True, hide_index=True,
                )
            else:
                st.info("Sin acciones registradas para este pozo.")
        else:
            st.info("No se pudo identificar la columna de pozo en la hoja Acciones.")
    else:
        st.info("Sin datos de acciones.")

    st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)

    # ══ ÚLTIMAS INTERVENCIONES ══
    st.markdown('<div class="section-title">Últimas Intervenciones</div>', unsafe_allow_html=True)

    ev_pozo = idx["eventos"].get(_psel_up, pd.DataFrame())
    if not ev_pozo.empty:
        cols_ev = [c for c in ["FECHA_INICIO", "TIPO_EVENTO", "OBSERVACION"] if c in ev_pozo.columns]
        rename_ev = {"FECHA_INICIO": "Fecha", "TIPO_EVENTO": "Tipo", "OBSERVACION": "Observación"}
        st.dataframe(
            ev_pozo[cols_ev].head(5).rename(columns=rename_ev).reset_index(drop=True),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("Sin intervenciones registradas para este pozo.")

    # ══ RECOMENDACIÓN OPERATIVA (al final) ══
    st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)
    rec_txt = str(row.get("RECOMENDACION") or "")
    if rec_txt and rec_txt not in ("–", "nan"):
        st.markdown(f"""
        <div class="rec-card">
          <div class="rec-card-label">Recomendación Operativa</div>
          <div class="rec-card-text">{rec_txt}</div>
        </div>""", unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="rec-card">
          <div class="rec-card-label">Recomendación Operativa</div>
          <div class="rec-card-text" style="color:#888; font-style:italic;">
            Ejecutar el script corregido para generar recomendaciones automáticas.
          </div>
        </div>""", unsafe_allow_html=True)

    # ── Footer ──
    st.markdown("<div style='margin-top:30px'></div>", unsafe_allow_html=True)
    st.markdown(
        f"<div style='text-align:center; color:#aaa; font-size:0.73rem;'>"
        f"SIDOP · Quintana Energy · {datetime.now().year}</div>",
        unsafe_allow_html=True
    )


# ── Invocar el fragment dentro del tab ──
with tab2:
    tab2_detalle(idx, maestro, df, last_update)
