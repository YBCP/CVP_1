"""
Sistema de Visitas Técnicas - CVP (Caja de Vivienda Popular)
Bogotá, Colombia
"""

import streamlit as st
import pandas as pd
import os
import io
from pathlib import Path
from datetime import date, datetime, time
import plotly.express as px
import plotly.graph_objects as go

# ─────────────────────────────────────────────
# GLOBAL CONFIG
# ─────────────────────────────────────────────
DATA_DIR     = Path(__file__).parent / "data"
FORMATOS_DIR = Path(__file__).parent / "formatos"
MAESTRO_PATH    = DATA_DIR / "maestro_predios_cvp.csv"
VISITAS_PATH    = DATA_DIR / "visitas.csv"
RESULTADOS_PATH = DATA_DIR / "resultados.csv"
TECNICOS_PATH   = DATA_DIR / "tecnicos.csv"
FICHA_TEMPLATE  = FORMATOS_DIR / "208-REAS-Ft-30 FICHA TECNICA DE RECONOCIMIENTO v7.xlsx"
INFORME_TEMPLATE= FORMATOS_DIR / "208-REAS-Ft-176 INFORME DE GESTION.docx"

COLOR_PRIMARY = "#003366"
COLOR_SECONDARY = "#005B96"
COLOR_SUCCESS = "#28a745"
COLOR_DANGER = "#dc3545"
COLOR_WARNING = "#ffc107"

CARGOS = [
    "Técnico Social",
    "Profesional Social",
    "Técnico Jurídico",
    "Profesional Jurídico",
    "Técnico Predial",
]

ESTADOS_VISITA = ["Pendiente", "Exitosa", "Fallida"]
OCUPACIONES = ["Ocupado", "Desocupado", "En construcción", "Lote vacío"]
TIPOS_CONSTRUCCION = [
    "Mampostería confinada",
    "Mampostería simple",
    "Madera",
    "Metálica",
    "Mixta",
    "Otro",
]
ESTADOS_CONSERVACION = ["Bueno", "Regular", "Malo", "Ruina"]
PISOS_OPTIONS = ["1", "2", "3", "4", "5", "6+"]
MOTIVOS_FALLIDA = [
    "Predio no encontrado",
    "Acceso negado",
    "Propietario no presente",
    "Predio demolido",
    "Dirección incorrecta",
    "Otro",
]

st.set_page_config(
    page_title="CVP - Sistema de Visitas",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown(
    f"""
    <style>
    /* Sidebar background */
    [data-testid="stSidebar"] {{
        background-color: {COLOR_PRIMARY};
    }}
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] div,
    [data-testid="stSidebar"] label {{
        color: white !important;
    }}
    /* Header strip */
    .cvp-header {{
        background-color: {COLOR_PRIMARY};
        color: white;
        padding: 18px 24px;
        border-radius: 8px;
        margin-bottom: 20px;
    }}
    .cvp-header h1 {{
        margin: 0;
        font-size: 1.6rem;
    }}
    .cvp-header p {{
        margin: 4px 0 0;
        font-size: 0.9rem;
        opacity: 0.85;
    }}
    /* Metric card */
    .metric-card {{
        background: white;
        border-left: 5px solid {COLOR_SECONDARY};
        border-radius: 6px;
        padding: 16px 20px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
        text-align: center;
    }}
    .metric-card .value {{
        font-size: 2.2rem;
        font-weight: 700;
        color: {COLOR_PRIMARY};
    }}
    .metric-card .label {{
        font-size: 0.85rem;
        color: #555;
        margin-top: 4px;
    }}
    /* Section title */
    .section-title {{
        color: {COLOR_PRIMARY};
        font-weight: 700;
        font-size: 1.1rem;
        border-bottom: 2px solid {COLOR_SECONDARY};
        padding-bottom: 4px;
        margin: 20px 0 12px;
    }}
    /* Result banner */
    .banner-exitosa {{
        background-color: {COLOR_SUCCESS};
        color: white;
        padding: 10px 20px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 1.1rem;
        text-align: center;
    }}
    .banner-fallida {{
        background-color: {COLOR_DANGER};
        color: white;
        padding: 10px 20px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 1.1rem;
        text-align: center;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────
# DATA LOADING HELPERS
# ─────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_maestro():
    try:
        df = pd.read_csv(MAESTRO_PATH, encoding="utf-8-sig", low_memory=False, dtype=str)
        df.columns = df.columns.str.strip()
        for col in ["LATITUD", "LONGITUD"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        if "AVALUO" in df.columns:
            df["AVALUO"] = pd.to_numeric(df["AVALUO"], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"Error cargando maestro: {e}")
        return pd.DataFrame()


def load_visitas():
    try:
        df = pd.read_csv(VISITAS_PATH, dtype=str)
        if df.empty or len(df.columns) == 0:
            return _empty_visitas()
        df.columns = df.columns.str.strip()
        return df
    except Exception:
        return _empty_visitas()


def _empty_visitas():
    return pd.DataFrame(
        columns=[
            "NUM_VISITA", "FECHA_PROGRAMADA", "REA", "SIN_REA",
            "DIRECCION_MANUAL", "LATITUD_MANUAL", "LONGITUD_MANUAL",
            "TECNICOS", "ESTADO", "NUM_VISITA_PREDIO", "FECHA_REGISTRO",
            "OBSERVACIONES_PROG",
        ]
    )


def load_resultados():
    try:
        df = pd.read_csv(RESULTADOS_PATH, dtype=str)
        if df.empty or len(df.columns) == 0:
            return _empty_resultados()
        df.columns = df.columns.str.strip()
        return df
    except Exception:
        return _empty_resultados()


def _empty_resultados():
    return pd.DataFrame(
        columns=[
            "NUM_VISITA", "REA", "FECHA_VISITA", "HORA_INICIO", "HORA_FIN",
            "TECNICOS", "RESULTADO", "OCUPACION", "PROP_CONTACTADO",
            "TIPO_CONSTRUCCION", "NUM_PISOS", "ESTADO_CONSERVACION",
            "LINDERO_NORTE", "LINDERO_SUR", "LINDERO_ORIENTE", "LINDERO_OCCIDENTE",
            "MOTIVO_FALLIDA", "OBSERVACIONES", "FOTOS", "FECHA_REGISTRO",
        ]
    )


def load_tecnicos():
    try:
        df = pd.read_csv(TECNICOS_PATH, dtype=str)
        df.columns = df.columns.str.strip()
        df["ACTIVO"] = df["ACTIVO"].str.strip().str.lower().map(
            {"true": True, "false": False, "1": True, "0": False}
        ).fillna(True)
        return df
    except Exception:
        return pd.DataFrame(
            columns=["ID_TECNICO", "NOMBRE", "CARGO", "EMAIL", "ACTIVO"]
        )


def save_visitas(df):
    df.to_csv(VISITAS_PATH, index=False)


def save_resultados(df):
    df.to_csv(RESULTADOS_PATH, index=False)


def save_tecnicos(df):
    df.to_csv(TECNICOS_PATH, index=False)


def next_num_visita(visitas_df):
    year = datetime.now().year
    prefix = f"VT-{year}-"
    existing = visitas_df["NUM_VISITA"].dropna() if "NUM_VISITA" in visitas_df.columns else pd.Series(dtype=str)
    year_nums = existing[existing.str.startswith(prefix, na=False)]
    if year_nums.empty:
        return f"{prefix}0001"
    maxn = year_nums.str.replace(prefix, "", regex=False).apply(
        lambda x: int(x) if x.isdigit() else 0
    ).max()
    return f"{prefix}{str(maxn + 1).zfill(4)}"


def count_visitas_predio(visitas_df, rea):
    if visitas_df.empty or "REA" not in visitas_df.columns:
        return 0
    return int((visitas_df["REA"] == str(rea)).sum())


# ─────────────────────────────────────────────
# SIDEBAR NAVIGATION
# ─────────────────────────────────────────────

with st.sidebar:
    st.markdown(
        f"""
        <div style='text-align:center; padding: 10px 0 20px;'>
            <div style='font-size:2.5rem;'>🏛️</div>
            <div style='font-size:1.1rem; font-weight:700; color:white;'>CVP</div>
            <div style='font-size:0.75rem; color:#aac4e0;'>Caja de Vivienda Popular<br>Bogotá</div>
        </div>
        <hr style='border-color: #336699; margin: 0 0 12px;'>
        """,
        unsafe_allow_html=True,
    )

    pagina = st.radio(
        "Navegación",
        options=[
            "🏠 Inicio",
            "📋 Programar Visita",
            "Visitas Programadas",
            "✅ Registrar Resultado",
            "📄 Descargar Formato",
            "Gestión de Técnicos",
            "📈 Indicadores",
        ],
        label_visibility="collapsed",
    )

    st.markdown(
        """
        <hr style='border-color: #336699; margin: 20px 0 10px;'>
        <div style='font-size:0.7rem; color:#aac4e0; text-align:center;'>
            Dirección de Reasentamientos<br>© 2025 CVP Bogotá
        </div>
        """,
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────
if "pdf_num_visita" not in st.session_state:
    st.session_state["pdf_num_visita"] = None
if "registro_num_visita" not in st.session_state:
    st.session_state["registro_num_visita"] = None


# ════════════════════════════════════════════════════════════════
# FORMAT GENERATION FUNCTIONS
# ════════════════════════════════════════════════════════════════

def _str(val):
    """Return clean string or empty."""
    if val is None or (isinstance(val, float) and str(val) == "nan"):
        return ""
    return str(val).strip()


def generar_ficha_tecnica(res_row, vis_row, maestro_row):
    """Pre-fill 208-REAS-Ft-30 FICHA TECNICA DE RECONOCIMIENTO v7.xlsx and return bytes."""
    import openpyxl
    from copy import copy

    wb = openpyxl.load_workbook(str(FICHA_TEMPLATE))
    ws = wb.active

    fecha_vis = _str(res_row.get("FECHA_VISITA", ""))
    rea       = _str(res_row.get("REA", ""))
    tecnicos  = _str(res_row.get("TECNICOS", "")).replace("|", " / ")
    localidad = _str(maestro_row.get("LOCALIDAD", "")) if not maestro_row.empty else ""
    barrio    = _str(maestro_row.get("BARRIO", ""))    if not maestro_row.empty else ""
    dir_campo = _str(maestro_row.get("DIRECCION", "")) if not maestro_row.empty else ""
    manzana   = _str(maestro_row.get("MANZANA", ""))   if not maestro_row.empty else ""
    lote      = _str(maestro_row.get("LOTE", ""))      if not maestro_row.empty else ""
    chip      = _str(maestro_row.get("CHIP", ""))      if not maestro_row.empty else ""
    lind_n    = _str(res_row.get("LINDERO_NORTE", ""))
    lind_s    = _str(res_row.get("LINDERO_SUR", ""))
    lind_or   = _str(res_row.get("LINDERO_ORIENTE", ""))
    lind_oc   = _str(res_row.get("LINDERO_OCCIDENTE", ""))

    # Cell mapping  (label row → value row, first cell of merged range)
    ws["A7"]  = fecha_vis       # FECHA DE ELABORACION
    ws["N7"]  = rea             # IDENTIFICADOR
    ws["A9"]  = tecnicos        # NOMBRE DE QUIEN ATIENDE
    ws["A12"] = localidad       # 2.1 LOCALIDAD
    ws["K12"] = barrio          # 2.2 BARRIO
    ws["A14"] = dir_campo       # 2.4 DIRECCION TOMADA EN CAMPO
    ws["A18"] = dir_campo       # 2.8 DIRECCIÓN CATASTRAL
    ws["S18"] = manzana         # MANZANA catastral
    ws["W18"] = lote            # LOTE catastral
    ws["A20"] = chip            # 2.9 CHIP CATASTRAL
    ws["F22"] = lind_n          # LINDERO NORTE
    ws["F23"] = lind_s          # LINDERO SUR
    ws["F24"] = lind_or         # LINDERO ORIENTE
    ws["F25"] = lind_oc         # LINDERO OCCIDENTE

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def generar_informe_gestion(res_row, vis_row, maestro_row):
    """Pre-fill 208-REAS-Ft-176 INFORME DE GESTION.docx and return bytes."""
    from docx import Document

    doc = Document(str(INFORME_TEMPLATE))

    rea        = _str(res_row.get("REA", ""))
    propietario= _str(maestro_row.get("PROPIETARIO_1", "")) if not maestro_row.empty else ""
    tecnicos   = _str(res_row.get("TECNICOS", "")).replace("|", " / ")
    fecha_vis  = _str(res_row.get("FECHA_VISITA", ""))
    motivo     = _str(res_row.get("MOTIVO_FALLIDA", "")).replace("|", ", ")
    observ     = _str(res_row.get("OBSERVACIONES", ""))

    # Table 0: datos básicos
    tabla = doc.tables[0]
    tabla.rows[0].cells[1].text = rea           # IDENTIFICADOR
    tabla.rows[1].cells[1].text = propietario   # BENEFICIARIO
    tabla.rows[4].cells[1].text = tecnicos      # NOMBRE PROFESIONAL
    tabla.rows[6].cells[1].text = fecha_vis     # FECHA

    # Escribir en el párrafo vacío SIGUIENTE a "Descripción Gestión realizada:" (P05)
    desc_text = f"Motivo de visita fallida: {motivo}"
    if observ:
        desc_text += "\n\nObservaciones: " + observ
    paragraphs = doc.paragraphs
    for idx, p in enumerate(paragraphs):
        if "Descripci" in p.text and "Gesti" in p.text:
            # El recuadro de texto está en los párrafos vacíos siguientes (P05 en adelante)
            siguiente = idx + 1
            if siguiente < len(paragraphs):
                target = paragraphs[siguiente]
                target.clear()
                target.add_run(desc_text)
            break

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ════════════════════════════════════════════════════════════════
# PAGE: INICIO
# ════════════════════════════════════════════════════════════════
if pagina == "🏠 Inicio":
    st.markdown(
        f"""
        <div class="cvp-header">
            <h1>🏛️ Sistema de Visitas Técnicas — CVP</h1>
            <p>Caja de Vivienda Popular · Dirección de Reasentamientos · Bogotá D.C.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    maestro = load_maestro()
    visitas = load_visitas()
    resultados = load_resultados()

    total_predios = len(maestro)
    total_programadas = len(visitas) if not visitas.empty else 0

    if not visitas.empty and "ESTADO" in visitas.columns:
        total_exitosas = int((visitas["ESTADO"] == "Exitosa").sum())
        total_fallidas = int((visitas["ESTADO"] == "Fallida").sum())
        total_pendientes = int((visitas["ESTADO"] == "Pendiente").sum())
    else:
        total_exitosas = total_fallidas = total_pendientes = 0

    total_realizadas = total_exitosas + total_fallidas

    col1, col2, col3, col4 = st.columns(4)
    cards = [
        (col1, total_predios, "Total Predios en Maestro", "🗂️"),
        (col2, total_programadas, "Visitas Programadas", "📋"),
        (col3, total_realizadas, "Visitas Realizadas", "✅"),
        (col4, total_pendientes, "Visitas Pendientes", "⏳"),
    ]
    for col, val, label, icon in cards:
        with col:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div style="font-size:1.8rem;">{icon}</div>
                    <div class="value">{val:,}</div>
                    <div class="label">{label}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">Acciones rápidas</div>', unsafe_allow_html=True)

    qcol1, qcol2, qcol3 = st.columns(3)
    with qcol1:
        st.info("📋 **Programar Visita**\nRegistre una nueva visita técnica a un predio del maestro.")
    with qcol2:
        st.info("✅ **Registrar Resultado**\nDocumente el resultado de una visita programada.")
    with qcol3:
        st.info("📄 **Generar PDF**\nDescargue el formato oficial de visita técnica.")

    if total_programadas > 0:
        st.markdown('<div class="section-title">Últimas visitas programadas</div>', unsafe_allow_html=True)
        show_cols = [c for c in ["NUM_VISITA", "FECHA_PROGRAMADA", "REA", "TECNICOS", "ESTADO"] if c in visitas.columns]
        st.dataframe(visitas[show_cols].tail(5), use_container_width=True, hide_index=True)


# ════════════════════════════════════════════════════════════════
# PAGE: PROGRAMAR VISITA
# ════════════════════════════════════════════════════════════════
elif pagina == "📋 Programar Visita":
    st.markdown(
        '<div class="cvp-header"><h1>Programar Visita Técnica</h1><p>Registre y programe visitas a predios del reasentamiento</p></div>',
        unsafe_allow_html=True,
    )

    maestro = load_maestro()
    tecnicos_df = load_tecnicos()
    tecnicos_activos = (
        tecnicos_df[tecnicos_df["ACTIVO"] == True]["NOMBRE"].tolist()
        if not tecnicos_df.empty else []
    )

    tab_rea, tab_sin_rea = st.tabs(["🔎 Con REA", "📍 Sin REA"])

    # ── TAB CON REA ──────────────────────────────────────────────
    with tab_rea:
        modo = st.radio("Modo de ingreso", ["Manual", "Carga masiva Excel"], horizontal=True)

        predios_seleccionados = []

        if modo == "Manual":
            # ── Inicializar lista acumulable en session_state ──────────
            if "lista_predios_manual" not in st.session_state:
                st.session_state["lista_predios_manual"] = []

            st.markdown('<div class="section-title">Agregar predios a la visita</div>', unsafe_allow_html=True)

            campo_busqueda = st.radio("Buscar por", ["REA", "CHIP"], horizontal=True, key="campo_busq_manual")

            if campo_busqueda == "REA":
                opciones = sorted(maestro["REA"].dropna().astype(str).str.strip().unique().tolist())
            else:
                opciones = sorted(maestro["CHIP"].dropna().astype(str).str.strip().unique().tolist()) if "CHIP" in maestro.columns else []

            col_sel, col_btn = st.columns([4, 1])
            with col_sel:
                valor_busqueda = st.selectbox(
                    f"Escriba o seleccione {campo_busqueda}",
                    options=[""] + opciones,
                    index=0,
                    placeholder=f"Empiece a escribir el {campo_busqueda}...",
                    help=f"Escriba los primeros dígitos para filtrar ({len(opciones):,} opciones)",
                    key="sel_busq_manual",
                )
            with col_btn:
                st.markdown("<br>", unsafe_allow_html=True)
                agregar = st.button("➕ Agregar", use_container_width=True)

            # ── Agregar predio a la lista ──────────────────────────────
            if agregar and valor_busqueda:
                rea_a_agregar = valor_busqueda.strip()
                if campo_busqueda == "CHIP":
                    match = maestro[maestro["CHIP"].str.strip() == rea_a_agregar] if "CHIP" in maestro.columns else pd.DataFrame()
                    rea_a_agregar = match["REA"].iloc[0] if not match.empty and pd.notna(match["REA"].iloc[0]) else rea_a_agregar
                if rea_a_agregar in st.session_state["lista_predios_manual"]:
                    st.warning(f"El predio **{rea_a_agregar}** ya está en la lista.")
                else:
                    st.session_state["lista_predios_manual"].append(rea_a_agregar)
                    st.success(f"✓ Predio **{rea_a_agregar}** agregado.")

            # ── Mostrar lista acumulada ────────────────────────────────
            lista_actual = st.session_state["lista_predios_manual"]
            if lista_actual:
                st.markdown(f'<div class="section-title">Predios en esta visita ({len(lista_actual)})</div>', unsafe_allow_html=True)

                show_cols = [c for c in ["REA", "CHIP", "CHIP_VALIDADO", "LOCALIDAD", "BARRIO", "DIRECCION", "ESTADO_REA"] if c in maestro.columns]
                df_lista = maestro[maestro["REA"].isin(lista_actual)][show_cols].copy()

                # Agregar predios sin REA en maestro (sin REA) si los hay
                reas_no_en_maestro = [r for r in lista_actual if r not in maestro["REA"].values]
                if reas_no_en_maestro:
                    extras = pd.DataFrame([{"REA": r} for r in reas_no_en_maestro])
                    df_lista = pd.concat([df_lista, extras], ignore_index=True)

                # Columna para eliminar
                df_lista[""] = "🗑️"
                st.dataframe(df_lista, use_container_width=True, hide_index=True)

                # Eliminar predio individual
                col_rm1, col_rm2 = st.columns([3, 1])
                with col_rm1:
                    predio_quitar = st.selectbox("Quitar predio de la lista", options=[""] + lista_actual, key="sel_quitar")
                with col_rm2:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🗑️ Quitar", use_container_width=True) and predio_quitar:
                        st.session_state["lista_predios_manual"].remove(predio_quitar)
                        st.rerun()

                if st.button("🧹 Limpiar toda la lista", use_container_width=False):
                    st.session_state["lista_predios_manual"] = []
                    st.rerun()

            predios_seleccionados = lista_actual
            st.session_state["predios_prog_manual"] = predios_seleccionados

        else:  # Carga masiva
            st.markdown('<div class="section-title">Carga masiva desde Excel</div>', unsafe_allow_html=True)
            archivo = st.file_uploader(
                "Suba un archivo Excel con columna REA",
                type=["xlsx", "xls"],
                help="El archivo debe tener al menos una columna llamada REA",
            )
            if archivo:
                try:
                    df_excel = pd.read_excel(archivo, dtype=str)
                    df_excel.columns = df_excel.columns.str.strip().str.upper()
                    if "REA" not in df_excel.columns:
                        st.error("El archivo no tiene columna REA")
                    else:
                        reas_excel = df_excel["REA"].dropna().str.strip().unique().tolist()
                        encontrados = maestro[maestro["REA"].isin(reas_excel)]
                        no_encontrados = [r for r in reas_excel if r not in maestro["REA"].values]

                        st.success(f"✓ {len(encontrados)} predios encontrados | ⚠️ {len(no_encontrados)} no encontrados")

                        if no_encontrados:
                            with st.expander(f"REAs no encontrados ({len(no_encontrados)})"):
                                st.write(no_encontrados)

                        show_cols = [c for c in ["REA", "CHIP", "LOCALIDAD", "BARRIO", "DIRECCION", "ESTADO_REA"] if c in encontrados.columns]
                        st.dataframe(encontrados[show_cols], use_container_width=True, hide_index=True)
                        predios_seleccionados = encontrados["REA"].tolist()
                        st.session_state["predios_prog_masiva"] = predios_seleccionados
                except Exception as e:
                    st.error(f"Error leyendo el archivo: {e}")

        # ── FORMULARIO DE PROGRAMACIÓN ─────────────────────────
        predios_a_programar = (
            predios_seleccionados
            or st.session_state.get("predios_prog_manual", [])
            or st.session_state.get("predios_prog_masiva", [])
        )

        if predios_a_programar:
            st.markdown('<div class="section-title">Datos de programación</div>', unsafe_allow_html=True)

            # ── Fecha / hora / observaciones (comunes) ─────────────────
            gcol1, gcol2 = st.columns(2)
            with gcol1:
                fecha_prog = st.date_input("Fecha programada", value=date.today(), key="fecha_prog_rea")
                hora_est   = st.time_input("Hora estimada", value=time(8, 0), key="hora_prog_rea")
            with gcol2:
                obs_prog = st.text_area("Observaciones de programación", height=100, key="obs_prog_rea")

            # ── Técnico por predio ─────────────────────────────────────
            st.markdown('<div class="section-title">Asignación de técnico por predio</div>', unsafe_allow_html=True)
            st.caption("Seleccione el técnico responsable de cada predio. Puede cambiar la asignación antes de confirmar.")

            # Inicializar asignaciones en session_state
            if "asig_tecnicos" not in st.session_state:
                st.session_state["asig_tecnicos"] = {}

            for rea in predios_a_programar:
                info = maestro[maestro["REA"] == rea]
                dir_label = info["DIRECCION"].iloc[0] if not info.empty and pd.notna(info["DIRECCION"].iloc[0]) else ""
                label = f"{rea}" + (f"  —  {dir_label}" if dir_label else "")
                prev = st.session_state["asig_tecnicos"].get(rea, [])
                tec_predio = st.multiselect(
                    label,
                    options=tecnicos_activos,
                    default=[t for t in prev if t in tecnicos_activos],
                    key=f"tec_predio_{rea}",
                )
                st.session_state["asig_tecnicos"][rea] = tec_predio

            if st.button("📅 Programar Visita(s)", use_container_width=True, key="btn_prog_rea"):
                sin_tecnico = [r for r in predios_a_programar if not st.session_state["asig_tecnicos"].get(r)]
                if sin_tecnico:
                    st.warning(f"Faltan técnicos en: {', '.join(sin_tecnico)}")
                else:
                    visitas_df = load_visitas()
                    nuevas = []
                    nums_asignados = []
                    for rea in predios_a_programar:
                        tecs = st.session_state["asig_tecnicos"].get(rea, [])
                        num_visita = next_num_visita(visitas_df)
                        num_predio = count_visitas_predio(visitas_df, rea) + 1
                        nueva = {
                            "NUM_VISITA": num_visita,
                            "FECHA_PROGRAMADA": str(fecha_prog),
                            "REA": str(rea),
                            "SIN_REA": "No",
                            "DIRECCION_MANUAL": "",
                            "LATITUD_MANUAL": "",
                            "LONGITUD_MANUAL": "",
                            "TECNICOS": "|".join(tecs),
                            "ESTADO": "Pendiente",
                            "NUM_VISITA_PREDIO": str(num_predio),
                            "FECHA_REGISTRO": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "OBSERVACIONES_PROG": obs_prog,
                        }
                        nuevas.append(nueva)
                        nums_asignados.append(num_visita)
                        visitas_df = pd.concat([visitas_df, pd.DataFrame([nueva])], ignore_index=True)

                    save_visitas(visitas_df)
                    st.session_state["predios_prog_manual"] = []
                    st.session_state["predios_prog_masiva"] = []
                    st.session_state["lista_predios_manual"] = []
                    st.session_state["asig_tecnicos"] = {}
                    st.success(
                        f"✅ {len(nuevas)} visita(s) programada(s) exitosamente.\n\n"
                        f"Números asignados: {', '.join(nums_asignados)}"
                    )

    # ── TAB SIN REA ──────────────────────────────────────────────
    with tab_sin_rea:
        st.markdown('<div class="section-title">Predio sin REA (visita de campo)</div>', unsafe_allow_html=True)

        with st.form("form_sin_rea"):
            scol1, scol2 = st.columns(2)
            with scol1:
                dir_manual = st.text_input("Dirección", placeholder="Ej: CL 50 # 12-34")
                barrio_manual = st.text_input("Barrio")
                localidad_manual = st.text_input("Localidad")
            with scol2:
                lat_manual = st.text_input("Latitud (opcional)", placeholder="4.6xxx")
                lon_manual = st.text_input("Longitud (opcional)", placeholder="-74.0xxx")
                if lat_manual and lon_manual:
                    st.markdown(
                        f"🗺️ [Ver en Maps](https://www.google.com/maps?q={lat_manual},{lon_manual})"
                    )

            st.markdown("---")
            sfcol1, sfcol2 = st.columns(2)
            with sfcol1:
                fecha_sin = st.date_input("Fecha programada", value=date.today(), key="fecha_sin")
                tecnicos_sin = st.multiselect(
                    "Técnicos asignados",
                    options=tecnicos_activos,
                    key="tec_sin",
                )
            with sfcol2:
                hora_sin = st.time_input("Hora estimada", value=time(8, 0), key="hora_sin")
                obs_sin = st.text_area("Observaciones", height=80, key="obs_sin")

            sub_sin = st.form_submit_button("📅 Programar Visita Sin REA", use_container_width=True)

        if sub_sin:
            if not dir_manual:
                st.warning("Ingrese al menos una dirección.")
            elif not tecnicos_sin:
                st.warning("Seleccione al menos un técnico.")
            else:
                visitas_df = load_visitas()
                num_visita = next_num_visita(visitas_df)
                nueva = {
                    "NUM_VISITA": num_visita,
                    "FECHA_PROGRAMADA": str(fecha_sin),
                    "REA": "",
                    "SIN_REA": "Si",
                    "DIRECCION_MANUAL": f"{dir_manual}, {barrio_manual}, {localidad_manual}".strip(", "),
                    "LATITUD_MANUAL": lat_manual,
                    "LONGITUD_MANUAL": lon_manual,
                    "TECNICOS": "|".join(tecnicos_sin),
                    "ESTADO": "Pendiente",
                    "NUM_VISITA_PREDIO": "1",
                    "FECHA_REGISTRO": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "OBSERVACIONES_PROG": obs_sin,
                }
                visitas_df = pd.concat([visitas_df, pd.DataFrame([nueva])], ignore_index=True)
                save_visitas(visitas_df)
                st.success(f"✅ Visita programada. Número asignado: **{num_visita}**")


# ════════════════════════════════════════════════════════════════
# PAGE: VISITAS PROGRAMADAS
# ════════════════════════════════════════════════════════════════
elif pagina == "Visitas Programadas":
    st.markdown(
        '<div class="cvp-header"><h1>Visitas Programadas</h1><p>Consulte y filtre el historial de visitas programadas</p></div>',
        unsafe_allow_html=True,
    )

    visitas = load_visitas()
    maestro = load_maestro()

    if visitas.empty:
        st.info("No hay visitas programadas aún. Vaya a 📋 Programar Visita para crear la primera.")
    else:
        # Merge with maestro
        if not maestro.empty and "REA" in visitas.columns:
            merge_cols = [
                c for c in ["REA", "LOCALIDAD", "BARRIO", "DIRECCION", "ESTADO_REA", "TIPO_PREDIO", "AVALUO"]
                if c in maestro.columns
            ]
            visitas_merge = visitas.merge(
                maestro[merge_cols].rename(columns={
                    "LOCALIDAD": "LOCALIDAD_M",
                    "BARRIO": "BARRIO_M",
                    "DIRECCION": "DIRECCION_M",
                }),
                on="REA",
                how="left",
            )
        else:
            visitas_merge = visitas.copy()

        # ── FILTROS ───────────────────────────────────────────
        st.markdown('<div class="section-title">Filtros</div>', unsafe_allow_html=True)
        fcol1, fcol2, fcol3, fcol4 = st.columns(4)

        with fcol1:
            estado_filtro = st.selectbox("Estado", ["Todos"] + ESTADOS_VISITA)
        with fcol2:
            tec_options = ["Todos"]
            if "TECNICOS" in visitas_merge.columns:
                all_tecs = set()
                for t in visitas_merge["TECNICOS"].dropna():
                    all_tecs.update(t.split("|"))
                tec_options += sorted(all_tecs)
            tec_filtro = st.selectbox("Técnico", tec_options)
        with fcol3:
            loc_options = ["Todas"]
            loc_col = "LOCALIDAD_M" if "LOCALIDAD_M" in visitas_merge.columns else None
            if loc_col:
                loc_options += sorted(visitas_merge[loc_col].dropna().unique().tolist())
            loc_filtro = st.selectbox("Localidad", loc_options)
        with fcol4:
            fecha_min_str = (
                visitas_merge["FECHA_PROGRAMADA"].min()
                if "FECHA_PROGRAMADA" in visitas_merge.columns and not visitas_merge["FECHA_PROGRAMADA"].isna().all()
                else str(date.today())
            )
            try:
                fecha_min = date.fromisoformat(str(fecha_min_str)[:10])
            except Exception:
                fecha_min = date.today()
            fechas = st.date_input(
                "Rango fechas",
                value=(fecha_min, date.today()),
                key="rango_fechas_prog",
            )

        # Apply filters
        df_filtrado = visitas_merge.copy()
        if estado_filtro != "Todos" and "ESTADO" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["ESTADO"] == estado_filtro]
        if tec_filtro != "Todos" and "TECNICOS" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["TECNICOS"].str.contains(tec_filtro, na=False)]
        if loc_filtro != "Todas" and loc_col and loc_col in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado[loc_col] == loc_filtro]
        if isinstance(fechas, (list, tuple)) and len(fechas) == 2 and "FECHA_PROGRAMADA" in df_filtrado.columns:
            df_filtrado["_fecha_dt"] = pd.to_datetime(df_filtrado["FECHA_PROGRAMADA"], errors="coerce")
            df_filtrado = df_filtrado[
                (df_filtrado["_fecha_dt"] >= pd.Timestamp(fechas[0]))
                & (df_filtrado["_fecha_dt"] <= pd.Timestamp(fechas[1]))
            ]
            df_filtrado = df_filtrado.drop(columns=["_fecha_dt"])

        # ── SUMMARY ───────────────────────────────────────────
        sm1, sm2, sm3, sm4 = st.columns(4)
        n_tot = len(df_filtrado)
        n_exit = int((df_filtrado["ESTADO"] == "Exitosa").sum()) if "ESTADO" in df_filtrado.columns else 0
        n_fall = int((df_filtrado["ESTADO"] == "Fallida").sum()) if "ESTADO" in df_filtrado.columns else 0
        n_pend = int((df_filtrado["ESTADO"] == "Pendiente").sum()) if "ESTADO" in df_filtrado.columns else 0

        sm1.metric("Total", n_tot)
        sm2.metric("Exitosas", n_exit)
        sm3.metric("Fallidas", n_fall)
        sm4.metric("Pendientes", n_pend)

        # ── COLOR CODING TABLE ────────────────────────────────
        show_cols = [
            c for c in [
                "NUM_VISITA", "FECHA_PROGRAMADA", "REA", "DIRECCION_MANUAL",
                "TECNICOS", "ESTADO", "NUM_VISITA_PREDIO",
                "LOCALIDAD_M", "BARRIO_M", "DIRECCION_M", "ESTADO_REA",
            ] if c in df_filtrado.columns
        ]
        display_df = df_filtrado[show_cols].copy()

        st.dataframe(display_df, use_container_width=True, hide_index=True)

        # ── EXPORT ────────────────────────────────────────────
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_filtrado.to_excel(writer, index=False, sheet_name="Visitas")
        st.download_button(
            "📥 Exportar a Excel",
            data=buf.getvalue(),
            file_name=f"visitas_cvp_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ════════════════════════════════════════════════════════════════
# PAGE: REGISTRAR RESULTADO
# ════════════════════════════════════════════════════════════════
elif pagina == "✅ Registrar Resultado":
    st.markdown(
        '<div class="cvp-header"><h1>Registrar Resultado de Visita</h1><p>Documente el resultado de una visita técnica programada</p></div>',
        unsafe_allow_html=True,
    )

    visitas = load_visitas()
    maestro = load_maestro()
    resultados = load_resultados()
    tecnicos_df = load_tecnicos()
    tecnicos_activos = (
        tecnicos_df[tecnicos_df["ACTIVO"] == True]["NOMBRE"].tolist()
        if not tecnicos_df.empty else []
    )

    if visitas.empty:
        st.info("No hay visitas programadas. Programe una visita primero.")
    else:
        pendientes = visitas[visitas["ESTADO"] == "Pendiente"] if "ESTADO" in visitas.columns else visitas

        if pendientes.empty:
            st.info("No hay visitas pendientes de registro.")
        else:
            # Build dropdown options
            def visita_label(row):
                rea = row.get("REA", "")
                dir_m = row.get("DIRECCION_MANUAL", "")
                nv = row.get("NUM_VISITA", "")
                desc = rea if rea else dir_m
                return f"{nv} | {desc}"

            opciones = {visita_label(row): row for _, row in pendientes.iterrows()}

            presel = st.session_state.get("registro_num_visita")
            presel_key = None
            if presel:
                for k, v in opciones.items():
                    if v.get("NUM_VISITA") == presel:
                        presel_key = k
                        break

            sel_label = st.selectbox(
                "Seleccione la visita a registrar",
                list(opciones.keys()),
                index=list(opciones.keys()).index(presel_key) if presel_key else 0,
            )
            visita_sel = opciones[sel_label]
            num_visita_sel = visita_sel.get("NUM_VISITA", "")
            rea_sel = visita_sel.get("REA", "")
            num_predio = visita_sel.get("NUM_VISITA_PREDIO", "1")

            # Predio info card
            if rea_sel and not maestro.empty:
                predio_info = maestro[maestro["REA"] == rea_sel]
                if not predio_info.empty:
                    pi = predio_info.iloc[0]
                    st.markdown('<div class="section-title">Información del Predio</div>', unsafe_allow_html=True)
                    ic1, ic2, ic3 = st.columns(3)
                    with ic1:
                        st.markdown(f"**REA:** {pi.get('REA','')}")
                        st.markdown(f"**CHIP:** {pi.get('CHIP','')}")
                        st.markdown(f"**Localidad:** {pi.get('LOCALIDAD','')}")
                    with ic2:
                        st.markdown(f"**Barrio:** {pi.get('BARRIO','')}")
                        st.markdown(f"**Dirección:** {pi.get('DIRECCION','')}")
                        st.markdown(f"**Estado REA:** {pi.get('ESTADO_REA','')}")
                    with ic3:
                        st.markdown(f"**Propietario:** {pi.get('PROPIETARIO_1','')}")
                        st.markdown(f"**Cédula:** {pi.get('CEDULA_1','')}")
                        avaluo = pi.get('AVALUO', 0)
                        try:
                            avaluo_fmt = f"${float(avaluo):,.0f}" if float(str(avaluo)) > 0 else "No registrado"
                        except Exception:
                            avaluo_fmt = str(avaluo)
                        st.markdown(f"**Avalúo:** {avaluo_fmt}")

            st.info(f"📌 Esta es la **visita N° {num_predio}** a este predio.")

            # ── RESULTADO FORM ────────────────────────────────
            st.markdown('<div class="section-title">Resultado de la Visita</div>', unsafe_allow_html=True)

            with st.form("form_resultado"):
                resultado = st.radio("Resultado", ["Exitosa", "Fallida"], horizontal=True)

                fecha_visita = st.date_input("Fecha de visita", value=date.today())

                # Técnicos programados como default, editables
                tecs_programados = [t.strip() for t in str(visita_sel.get("TECNICOS", "")).split("|") if t.strip()]
                tecs_validos = [t for t in tecs_programados if t in tecnicos_activos]
                tecs_resultado = st.multiselect(
                    "Técnico(s) que realizaron la visita",
                    options=tecnicos_activos,
                    default=tecs_validos,
                    help="Pre-cargado con los técnicos programados. Modifique si hubo cambios.",
                )
                tec_display = "|".join(tecs_resultado)

                if resultado == "Exitosa":
                    rc1, rc2 = st.columns(2)
                    with rc1:
                        hora_ini = st.time_input("Hora inicio", value=time(9, 0))
                        ocupacion = st.selectbox("Ocupación", OCUPACIONES)
                        tipo_const = st.selectbox("Tipo de construcción", TIPOS_CONSTRUCCION)
                        num_pisos = st.selectbox("Número de pisos", PISOS_OPTIONS)
                    with rc2:
                        hora_fin = st.time_input("Hora fin", value=time(10, 0))
                        prop_contactado = st.radio("Propietario contactado", ["Si", "No"], horizontal=True)
                        estado_cons = st.selectbox("Estado de conservación", ESTADOS_CONSERVACION)

                    st.markdown("**Linderos**")
                    lc1, lc2, lc3, lc4 = st.columns(4)
                    with lc1:
                        lindero_n = st.text_input("Norte", key="ln")
                    with lc2:
                        lindero_s = st.text_input("Sur", key="ls")
                    with lc3:
                        lindero_o = st.text_input("Oriente", key="lor")
                    with lc4:
                        lindero_oc = st.text_input("Occidente", key="loc")

                    observaciones = st.text_area("Observaciones", height=100)

                    motivo_fallida = ""
                    hora_ini_f = str(hora_ini)
                    hora_fin_f = str(hora_fin)

                else:  # Fallida
                    st.markdown("**Motivo de visita fallida**")
                    motivos_sel = []
                    mf_cols = st.columns(2)
                    for i, mot in enumerate(MOTIVOS_FALLIDA):
                        with mf_cols[i % 2]:
                            if st.checkbox(mot, key=f"mot_{i}"):
                                motivos_sel.append(mot)

                    observaciones = st.text_area("Observaciones", height=100, key="obs_fallida")

                    motivo_fallida = "|".join(motivos_sel)
                    hora_ini_f = hora_fin_f = ""
                    ocupacion = prop_contactado = tipo_const = num_pisos = estado_cons = ""
                    lindero_n = lindero_s = lindero_o = lindero_oc = ""

                sub_res = st.form_submit_button("💾 Guardar Resultado", use_container_width=True)

            if sub_res:
                # Build result row
                nuevo_resultado = {
                    "NUM_VISITA": num_visita_sel,
                    "REA": rea_sel,
                    "FECHA_VISITA": str(fecha_visita),
                    "HORA_INICIO": hora_ini_f if resultado == "Exitosa" else "",
                    "HORA_FIN": hora_fin_f if resultado == "Exitosa" else "",
                    "TECNICOS": tec_display,
                    "RESULTADO": resultado,
                    "OCUPACION": ocupacion if resultado == "Exitosa" else "",
                    "PROP_CONTACTADO": prop_contactado if resultado == "Exitosa" else "",
                    "TIPO_CONSTRUCCION": tipo_const if resultado == "Exitosa" else "",
                    "NUM_PISOS": num_pisos if resultado == "Exitosa" else "",
                    "ESTADO_CONSERVACION": estado_cons if resultado == "Exitosa" else "",
                    "LINDERO_NORTE": lindero_n if resultado == "Exitosa" else "",
                    "LINDERO_SUR": lindero_s if resultado == "Exitosa" else "",
                    "LINDERO_ORIENTE": lindero_o if resultado == "Exitosa" else "",
                    "LINDERO_OCCIDENTE": lindero_oc if resultado == "Exitosa" else "",
                    "MOTIVO_FALLIDA": motivo_fallida if resultado == "Fallida" else "",
                    "OBSERVACIONES": observaciones,
                    "FOTOS": "",
                    "FECHA_REGISTRO": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }

                resultados_df = load_resultados()
                resultados_df = pd.concat(
                    [resultados_df, pd.DataFrame([nuevo_resultado])], ignore_index=True
                )
                save_resultados(resultados_df)

                # Update visitas estado
                visitas_df = load_visitas()
                visitas_df.loc[visitas_df["NUM_VISITA"] == num_visita_sel, "ESTADO"] = resultado
                save_visitas(visitas_df)

                st.success(f"✅ Resultado '{resultado}' guardado para visita **{num_visita_sel}**")
                st.session_state["pdf_num_visita"] = num_visita_sel
                st.session_state["registro_num_visita"] = None
                st.info("Vaya a **📄 Descargar Formato** en el menú lateral para obtener el formato pre-diligenciado.")


# ════════════════════════════════════════════════════════════════
# PAGE: GENERAR FORMATO PDF
# ════════════════════════════════════════════════════════════════
# ════════════════════════════════════════════════════════════════
# PAGE: DESCARGAR FORMATO
# ════════════════════════════════════════════════════════════════
elif pagina == "📄 Descargar Formato":
    st.markdown(
        '<div class="cvp-header"><h1>Descargar Formato Oficial</h1>'
        '<p>Descargue el formato pre-diligenciado según el resultado de la visita</p></div>',
        unsafe_allow_html=True,
    )

    resultados = load_resultados()
    visitas    = load_visitas()
    maestro    = load_maestro()

    if resultados.empty:
        st.info("No hay resultados registrados aún. Primero registre el resultado de una visita.")
    else:
        presel_pdf   = st.session_state.get("pdf_num_visita")
        options_pdf  = resultados["NUM_VISITA"].dropna().tolist()

        if not options_pdf:
            st.info("No hay visitas con resultado registrado.")
        else:
            idx_presel = 0
            if presel_pdf and presel_pdf in options_pdf:
                idx_presel = options_pdf.index(presel_pdf)

            num_visita_pdf = st.selectbox("Seleccione la visita", options_pdf, index=idx_presel)

            res_row  = resultados[resultados["NUM_VISITA"] == num_visita_pdf].iloc[0]
            vis_match = visitas[visitas["NUM_VISITA"] == num_visita_pdf]
            vis_row  = vis_match.iloc[0] if not vis_match.empty else pd.Series()

            rea_pdf      = _str(res_row.get("REA", ""))
            maestro_row  = pd.Series()
            if rea_pdf and not maestro.empty:
                m_match = maestro[maestro["REA"] == rea_pdf]
                if not m_match.empty:
                    maestro_row = m_match.iloc[0]

            resultado_val = _str(res_row.get("RESULTADO", ""))

            # ── Vista previa ──────────────────────────────────────────────
            st.markdown('<div class="section-title">Datos de la visita</div>', unsafe_allow_html=True)
            if resultado_val == "Exitosa":
                st.markdown('<div class="banner-exitosa">✅ VISITA EXITOSA — Formato: Ficha Técnica de Reconocimiento</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="banner-fallida">❌ VISITA FALLIDA — Formato: Informe de Gestión</div>', unsafe_allow_html=True)

            pv1, pv2, pv3 = st.columns(3)
            with pv1:
                st.markdown(f"**N° Visita:** {num_visita_pdf}")
                st.markdown(f"**Fecha:** {_str(res_row.get('FECHA_VISITA',''))}")
                st.markdown(f"**REA:** {rea_pdf or '—'}")
            with pv2:
                st.markdown(f"**Técnico(s):** {_str(res_row.get('TECNICOS','')).replace('|', ', ')}")
                st.markdown(f"**Resultado:** {resultado_val}")
                num_vis_predio = _str(vis_row.get('NUM_VISITA_PREDIO','')) if not vis_row.empty else ''
                st.markdown(f"**Visita N° al predio:** {num_vis_predio or '—'}")
            with pv3:
                st.markdown(f"**Dirección:** {_str(maestro_row.get('DIRECCION','')) if not maestro_row.empty else '—'}")
                st.markdown(f"**Localidad:** {_str(maestro_row.get('LOCALIDAD','')) if not maestro_row.empty else '—'}")
                st.markdown(f"**Barrio:** {_str(maestro_row.get('BARRIO','')) if not maestro_row.empty else '—'}")

            st.markdown("---")

            # ── Botón de descarga según resultado ─────────────────────────
            if resultado_val == "Exitosa":
                if not FICHA_TEMPLATE.exists():
                    st.error(f"No se encontró la plantilla: {FICHA_TEMPLATE}")
                else:
                    if st.button("📥 Generar Ficha Técnica (XLSX)", use_container_width=True):
                        xlsx_bytes = generar_ficha_tecnica(res_row, vis_row, maestro_row)
                        st.download_button(
                            label="⬇️ Descargar 208-REAS-Ft-30 Ficha Técnica",
                            data=xlsx_bytes,
                            file_name=f"FichaTecnica_{num_visita_pdf}_{rea_pdf}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                    st.caption("El archivo descargado tiene los datos del predio pre-diligenciados. Complete los campos restantes en Excel.")
            else:
                if not INFORME_TEMPLATE.exists():
                    st.error(f"No se encontró la plantilla: {INFORME_TEMPLATE}")
                else:
                    if st.button("📥 Generar Informe de Gestión (DOCX)", use_container_width=True):
                        docx_bytes = generar_informe_gestion(res_row, vis_row, maestro_row)
                        st.download_button(
                            label="⬇️ Descargar 208-REAS-Ft-176 Informe de Gestión",
                            data=docx_bytes,
                            file_name=f"InformeGestion_{num_visita_pdf}_{rea_pdf}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                        )
                    st.caption("El archivo descargado tiene los datos básicos pre-diligenciados. Complete los campos restantes en Word.")


# ════════════════════════════════════════════════════════════════
# PAGE: GESTIÓN DE TÉCNICOS
# ════════════════════════════════════════════════════════════════
elif pagina == "Gestión de Técnicos":
    st.markdown(
        '<div class="cvp-header"><h1>Gestión de Técnicos</h1><p>Administre el equipo técnico de reasentamientos</p></div>',
        unsafe_allow_html=True,
    )

    tecnicos_df = load_tecnicos()

    st.markdown('<div class="section-title">Técnicos registrados</div>', unsafe_allow_html=True)

    if not tecnicos_df.empty:
        edited = st.data_editor(
            tecnicos_df,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "ID_TECNICO": st.column_config.TextColumn("ID", disabled=True),
                "NOMBRE": st.column_config.TextColumn("Nombre completo"),
                "CARGO": st.column_config.SelectboxColumn("Cargo", options=CARGOS),
                "EMAIL": st.column_config.TextColumn("Correo electrónico"),
                "ACTIVO": st.column_config.CheckboxColumn("Activo"),
            },
            hide_index=True,
        )

        if st.button("💾 Guardar cambios", key="save_tec_edit"):
            save_tecnicos(edited)
            st.success("✅ Cambios guardados exitosamente.")
            st.rerun()

    st.markdown('<div class="section-title">Agregar nuevo técnico</div>', unsafe_allow_html=True)

    with st.form("form_nuevo_tecnico"):
        nc1, nc2 = st.columns(2)
        with nc1:
            nuevo_nombre = st.text_input("Nombre completo")
            nuevo_cargo = st.selectbox("Cargo", CARGOS)
        with nc2:
            nuevo_email = st.text_input("Correo electrónico")
            nuevo_activo = st.checkbox("Activo", value=True)

        sub_tec = st.form_submit_button("➕ Agregar Técnico", use_container_width=True)

    if sub_tec:
        if not nuevo_nombre:
            st.warning("Ingrese el nombre del técnico.")
        else:
            tecnicos_df = load_tecnicos()
            # Generate new ID
            if tecnicos_df.empty:
                new_id = "T001"
            else:
                ids = tecnicos_df["ID_TECNICO"].str.replace("T", "", regex=False)
                max_id = ids.apply(lambda x: int(x) if x.isdigit() else 0).max()
                new_id = f"T{str(max_id + 1).zfill(3)}"

            nuevo_tecnico = {
                "ID_TECNICO": new_id,
                "NOMBRE": nuevo_nombre,
                "CARGO": nuevo_cargo,
                "EMAIL": nuevo_email,
                "ACTIVO": nuevo_activo,
            }
            tecnicos_df = pd.concat(
                [tecnicos_df, pd.DataFrame([nuevo_tecnico])], ignore_index=True
            )
            save_tecnicos(tecnicos_df)
            st.success(f"✅ Técnico '{nuevo_nombre}' agregado con ID {new_id}.")
            st.rerun()


# ════════════════════════════════════════════════════════════════
# PAGE: INDICADORES
# ════════════════════════════════════════════════════════════════
elif pagina == "📈 Indicadores":
    st.markdown(
        '<div class="cvp-header"><h1>Indicadores de Gestión</h1><p>Análisis y estadísticas del programa de visitas técnicas</p></div>',
        unsafe_allow_html=True,
    )

    visitas = load_visitas()
    resultados = load_resultados()
    tecnicos_df = load_tecnicos()

    tab_gen, tab_tec, tab_pred = st.tabs(["📊 General", "👤 Por Técnico", "🏘️ Por Predio"])

    # ── TAB GENERAL ───────────────────────────────────────────
    with tab_gen:
        if visitas.empty:
            st.info("No hay datos de visitas aún.")
        else:
            # Date filter
            gf1, gf2 = st.columns([2, 2])
            with gf1:
                periodo = st.selectbox("Período", ["Esta semana", "Este mes", "Rango personalizado"])
            with gf2:
                if periodo == "Rango personalizado":
                    rango_g = st.date_input("Rango", value=(date.today().replace(day=1), date.today()), key="rango_g")
                else:
                    rango_g = None

            vis_g = visitas.copy()
            if "FECHA_PROGRAMADA" in vis_g.columns:
                vis_g["_fecha"] = pd.to_datetime(vis_g["FECHA_PROGRAMADA"], errors="coerce")
                today = pd.Timestamp(date.today())
                if periodo == "Esta semana":
                    start = today - pd.Timedelta(days=today.dayofweek)
                    vis_g = vis_g[vis_g["_fecha"] >= start]
                elif periodo == "Este mes":
                    vis_g = vis_g[vis_g["_fecha"].dt.month == today.month]
                elif rango_g and isinstance(rango_g, (list, tuple)) and len(rango_g) == 2:
                    vis_g = vis_g[
                        (vis_g["_fecha"] >= pd.Timestamp(rango_g[0]))
                        & (vis_g["_fecha"] <= pd.Timestamp(rango_g[1]))
                    ]

            tot_g = len(vis_g)
            exit_g = int((vis_g["ESTADO"] == "Exitosa").sum()) if "ESTADO" in vis_g.columns else 0
            fall_g = int((vis_g["ESTADO"] == "Fallida").sum()) if "ESTADO" in vis_g.columns else 0
            pct_g = round(exit_g / tot_g * 100, 1) if tot_g > 0 else 0

            mc1, mc2, mc3, mc4 = st.columns(4)
            mc1.metric("Total visitas", tot_g)
            mc2.metric("Exitosas", exit_g)
            mc3.metric("Fallidas", fall_g)
            mc4.metric("% Éxito", f"{pct_g}%")

            if not vis_g.empty and "ESTADO" in vis_g.columns and "_fecha" in vis_g.columns:
                vis_g["_semana"] = vis_g["_fecha"].dt.to_period("W").apply(lambda r: str(r.start_time.date()) if pd.notna(r) else None)
                sem_data = vis_g.groupby(["_semana", "ESTADO"]).size().reset_index(name="count")

                if not sem_data.empty:
                    gc1, gc2 = st.columns(2)
                    with gc1:
                        fig_bar = px.bar(
                            sem_data,
                            x="_semana",
                            y="count",
                            color="ESTADO",
                            color_discrete_map={"Exitosa": COLOR_SUCCESS, "Fallida": COLOR_DANGER, "Pendiente": COLOR_WARNING},
                            title="Visitas por semana",
                            labels={"_semana": "Semana", "count": "Visitas"},
                            barmode="stack",
                        )
                        fig_bar.update_layout(plot_bgcolor="white", paper_bgcolor="white")
                        st.plotly_chart(fig_bar, use_container_width=True)

                    with gc2:
                        pie_data = vis_g["ESTADO"].value_counts().reset_index()
                        pie_data.columns = ["ESTADO", "count"]
                        fig_pie = px.pie(
                            pie_data,
                            names="ESTADO",
                            values="count",
                            color="ESTADO",
                            color_discrete_map={"Exitosa": COLOR_SUCCESS, "Fallida": COLOR_DANGER, "Pendiente": COLOR_WARNING},
                            title="Distribución por resultado",
                        )
                        st.plotly_chart(fig_pie, use_container_width=True)

            # Bar chart by localidad
            maestro = load_maestro()
            if not vis_g.empty and "REA" in vis_g.columns and not maestro.empty:
                merge_loc = vis_g.merge(
                    maestro[["REA", "LOCALIDAD"]].drop_duplicates(),
                    on="REA",
                    how="left",
                )
                loc_data = merge_loc.groupby(["LOCALIDAD", "ESTADO"]).size().reset_index(name="count") if "ESTADO" in merge_loc.columns else pd.DataFrame()
                if not loc_data.empty and loc_data["LOCALIDAD"].notna().any():
                    fig_loc = px.bar(
                        loc_data.dropna(subset=["LOCALIDAD"]),
                        x="LOCALIDAD",
                        y="count",
                        color="ESTADO",
                        color_discrete_map={"Exitosa": COLOR_SUCCESS, "Fallida": COLOR_DANGER, "Pendiente": COLOR_WARNING},
                        title="Visitas por localidad",
                        barmode="stack",
                    )
                    fig_loc.update_layout(xaxis_tickangle=-45, plot_bgcolor="white", paper_bgcolor="white")
                    st.plotly_chart(fig_loc, use_container_width=True)

    # ── TAB POR TÉCNICO ───────────────────────────────────────
    with tab_tec:
        if visitas.empty or "TECNICOS" not in visitas.columns:
            st.info("No hay datos suficientes.")
        else:
            # Explode technicians
            vis_tec = visitas.copy()
            vis_tec["_fecha"] = pd.to_datetime(vis_tec.get("FECHA_PROGRAMADA"), errors="coerce")
            vis_tec_exp = vis_tec.assign(
                TECNICO=vis_tec["TECNICOS"].str.split("|")
            ).explode("TECNICO")
            vis_tec_exp["TECNICO"] = vis_tec_exp["TECNICO"].str.strip()

            resumen_tec = vis_tec_exp.groupby("TECNICO").agg(
                Total=("NUM_VISITA", "count"),
                Exitosas=("ESTADO", lambda x: (x == "Exitosa").sum()),
                Fallidas=("ESTADO", lambda x: (x == "Fallida").sum()),
            ).reset_index()
            resumen_tec["% Éxito"] = (resumen_tec["Exitosas"] / resumen_tec["Total"] * 100).round(1)

            st.markdown('<div class="section-title">Resumen por técnico</div>', unsafe_allow_html=True)
            st.dataframe(resumen_tec, use_container_width=True, hide_index=True)

            fig_tec = px.bar(
                resumen_tec,
                x="TECNICO",
                y=["Exitosas", "Fallidas"],
                title="Visitas por técnico",
                barmode="stack",
                color_discrete_map={"Exitosas": COLOR_SUCCESS, "Fallidas": COLOR_DANGER},
            )
            fig_tec.update_layout(xaxis_tickangle=-30, plot_bgcolor="white", paper_bgcolor="white")
            st.plotly_chart(fig_tec, use_container_width=True)

            # Weekly evolution
            vis_tec_exp["_semana"] = vis_tec_exp["_fecha"].dt.to_period("W").apply(
                lambda r: str(r.start_time.date()) if pd.notna(r) else None
            )
            sem_tec = vis_tec_exp.groupby(["_semana", "TECNICO"]).size().reset_index(name="count")
            if not sem_tec.empty and sem_tec["_semana"].notna().any():
                fig_line = px.line(
                    sem_tec.dropna(subset=["_semana"]),
                    x="_semana",
                    y="count",
                    color="TECNICO",
                    title="Evolución semanal por técnico",
                    markers=True,
                    labels={"_semana": "Semana", "count": "Visitas"},
                )
                fig_line.update_layout(plot_bgcolor="white", paper_bgcolor="white")
                st.plotly_chart(fig_line, use_container_width=True)

    # ── TAB POR PREDIO ────────────────────────────────────────
    with tab_pred:
        st.markdown('<div class="section-title">Historial por predio</div>', unsafe_allow_html=True)
        rea_busq = st.text_input("Ingrese REA del predio")

        if rea_busq:
            vis_predio = visitas[visitas["REA"] == rea_busq.strip()] if "REA" in visitas.columns else pd.DataFrame()
            res_predio = resultados[resultados["REA"] == rea_busq.strip()] if "REA" in resultados.columns else pd.DataFrame()

            if vis_predio.empty:
                st.warning(f"No se encontraron visitas para REA '{rea_busq}'")
            else:
                maestro = load_maestro()
                if not maestro.empty:
                    predio_m = maestro[maestro["REA"] == rea_busq.strip()]
                    if not predio_m.empty:
                        pi = predio_m.iloc[0]
                        st.markdown(
                            f"""
                            **Predio:** {pi.get('REA','')} | **Dirección:** {pi.get('DIRECCION','')} |
                            **Barrio:** {pi.get('BARRIO','')} | **Localidad:** {pi.get('LOCALIDAD','')} |
                            **Estado REA:** {pi.get('ESTADO_REA','')}
                            """
                        )

                st.markdown(f"**Total visitas realizadas:** {len(vis_predio)}")

                show_vp = [c for c in ["NUM_VISITA", "FECHA_PROGRAMADA", "TECNICOS", "ESTADO", "NUM_VISITA_PREDIO", "OBSERVACIONES_PROG"] if c in vis_predio.columns]
                st.dataframe(vis_predio[show_vp], use_container_width=True, hide_index=True)

                if not res_predio.empty:
                    st.markdown('<div class="section-title">Resultados registrados</div>', unsafe_allow_html=True)
                    show_rp = [c for c in ["NUM_VISITA", "FECHA_VISITA", "RESULTADO", "OCUPACION", "TIPO_CONSTRUCCION", "OBSERVACIONES"] if c in res_predio.columns]
                    st.dataframe(res_predio[show_rp], use_container_width=True, hide_index=True)

                # Timeline
                if "FECHA_PROGRAMADA" in vis_predio.columns and "ESTADO" in vis_predio.columns:
                    st.markdown('<div class="section-title">Línea de tiempo</div>', unsafe_allow_html=True)
                    tl_df = vis_predio[["NUM_VISITA", "FECHA_PROGRAMADA", "ESTADO", "TECNICOS"]].copy()
                    tl_df["_fecha"] = pd.to_datetime(tl_df["FECHA_PROGRAMADA"], errors="coerce")
                    tl_df = tl_df.dropna(subset=["_fecha"])
                    if not tl_df.empty:
                        fig_tl = px.scatter(
                            tl_df,
                            x="_fecha",
                            y=[rea_busq] * len(tl_df),
                            color="ESTADO",
                            color_discrete_map={"Exitosa": COLOR_SUCCESS, "Fallida": COLOR_DANGER, "Pendiente": COLOR_WARNING},
                            hover_data=["NUM_VISITA", "TECNICOS"],
                            title=f"Línea de tiempo - REA {rea_busq}",
                            size_max=18,
                            size=[20] * len(tl_df),
                        )
                        fig_tl.update_layout(
                            yaxis_visible=False,
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                        )
                        st.plotly_chart(fig_tl, use_container_width=True)

