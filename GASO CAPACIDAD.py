import io
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.units import inch


# =========================
# STREAMLIT CONFIG
# =========================
st.set_page_config(page_title="GASO | Ocupación Crossdock", layout="wide")


# =========================
# CONFIG: m² por tipo pallet
# =========================
PALLET_M2_BY_TYPE = {
    "SOBREDIMENSIONADO": 3.6,
    "ESTANDAR": 1.44,
    "ESTÁNDAR": 1.44,
    "EUROPALLET": 0.96,
    "GABINETE GD": 1.44,
    "GABINETE MED": 1.44,
    "ANTENA CH": 2.88,
    "ANTENA MD": 2.88,
    "ANTENA GD": 3.6,
    "GALVANIZADO": 3.6,
    "CAJA": 0.96,
    "SOPORTE": 3.6,
}

# =========================
# CONFIG: capacidades m² por XDOCK
# =========================
PRESET_XDOCK_CAPACITY_M2 = {
    "Gaso- Tijuana-E-NS": 677.0,
    "Gaso- La Paz-E-NS": 316.0,
    "Gaso- Hermosillo-E-NS": 400.0,
    "Gaso- Culiacán-E-NS": 600.0,
    "Gaso- Cd. Juarez-E-NS": 350.0,
    "Gaso- Chihuahua-E-NS": 400.0,
    "Gaso- Monterrey-E-NS": 1500.0,
    "Gaso- Guadalajara-E-NS": 1200.0,
    "Gaso- Querétaro-E-NS": 800.0,
}

DEFAULT_TARGET_CARRIERS = ["TELCEL", "AT&T"]

# Nombres canónicos internos
COL_CARRIER = "CARRIER"
COL_XDOCK = "XDOCK"
COL_TIPO = "TIPO DE PALLET"
COL_ESTATUS = "ESTATUS DE SALIDA"


# =========================
# HELPERS
# =========================
def canon(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.upper()
    s = " ".join(s.split())
    return s


def normalize_and_map_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    alias_map = {
        canon("CARRIER"): COL_CARRIER,
        canon("XDOCK"): COL_XDOCK,
        canon("TIPO DE PALLET"): COL_TIPO,
        canon("TIPO PALLET"): COL_TIPO,
        canon("ESTATUS DE SALIDA"): COL_ESTATUS,
        canon("ESTATUS SALIDA"): COL_ESTATUS,
        canon("STATUS DE SALIDA"): COL_ESTATUS,
        canon("STATUS SALIDA"): COL_ESTATUS,
    }

    rename_dict = {}
    for c in df.columns:
        k = canon(c)
        if k in alias_map:
            rename_dict[c] = alias_map[k]

    df = df.rename(columns=rename_dict)
    return df


def ensure_required_columns(df: pd.DataFrame) -> list:
    req = [COL_CARRIER, COL_XDOCK, COL_TIPO, COL_ESTATUS]
    return [c for c in req if c not in df.columns]


def is_in_inventory(row: pd.Series) -> bool:
    # Regla: si ESTATUS DE SALIDA == "SALIDA" => NO contar
    estatus = row.get(COL_ESTATUS)
    estatus_txt = "" if pd.isna(estatus) else str(estatus).strip().upper()
    return estatus_txt != "SALIDA"


def compute_row_m2(row: pd.Series) -> float:
    # 1 fila = 1 pallet => m² depende SOLO del tipo
    t = str(row.get(COL_TIPO, "")).strip().upper()
    t = " ".join(t.split())
    if t not in PALLET_M2_BY_TYPE:
        return float("nan")
    return float(PALLET_M2_BY_TYPE[t])


def build_active(df_raw: pd.DataFrame, carriers_filter: list | None) -> pd.DataFrame:
    df = normalize_and_map_columns(df_raw)

    missing = ensure_required_columns(df)
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}. Columnas detectadas: {list(df.columns)}")

    # Normaliza campos clave
    df[COL_CARRIER] = df[COL_CARRIER].astype(str).str.strip()
    df[COL_XDOCK] = df[COL_XDOCK].astype(str).str.strip()
    df[COL_TIPO] = (
        df[COL_TIPO].astype(str).str.strip().str.upper().str.replace(r"\s+", " ", regex=True)
    )

    # Activos según ESTATUS != SALIDA
    active = df[df.apply(is_in_inventory, axis=1)].copy()

    # Filtra carriers
    if carriers_filter and len(carriers_filter) > 0 and "TODOS" not in carriers_filter:
        wanted = {c.upper() for c in carriers_filter}
        active = active[active[COL_CARRIER].str.upper().isin(wanted)].copy()

    # 1 fila = 1 pallet
    active["PALLETS_FILA"] = 1
    active["M2_OCUPADOS_FILA"] = active.apply(compute_row_m2, axis=1)

    return active


def summarize(active: pd.DataFrame, capacity_map: dict):
    active = active.copy()
    active["CAPACIDAD_M2_XDOCK"] = active[COL_XDOCK].map(capacity_map)

    # Por XDOCK (agregado)
    by_xdock = (
        active.groupby([COL_XDOCK], dropna=False)
        .agg(
            pallets=("PALLETS_FILA", "sum"),
            m2_ocupados=("M2_OCUPADOS_FILA", "sum"),
            capacidad_m2=("CAPACIDAD_M2_XDOCK", "first"),
        )
        .reset_index()
    )

    def pct(row):
        cap = row["capacidad_m2"]
        occ = row["m2_ocupados"]
        if pd.isna(cap) or cap == 0 or pd.isna(occ):
            return float("nan")
        return (occ / cap) * 100.0

    by_xdock["ocupacion_%"] = by_xdock.apply(pct, axis=1)
    by_xdock = by_xdock.sort_values("ocupacion_%", ascending=False, na_position="last").reset_index(drop=True)

    # Por Carrier + XDOCK
    by_carrier_xdock = (
        active.groupby([COL_CARRIER, COL_XDOCK], dropna=False)
        .agg(
            pallets=("PALLETS_FILA", "sum"),
            m2_ocupados=("M2_OCUPADOS_FILA", "sum"),
        )
        .reset_index()
        .sort_values([COL_CARRIER, "m2_ocupados"], ascending=[True, False])
        .reset_index(drop=True)
    )

    # Conteo por Carrier + XDOCK + Tipo (para reporte ejecutivo)
    cxt = (
        active.groupby([COL_CARRIER, COL_XDOCK, COL_TIPO], dropna=False)
        .size()
        .reset_index(name="pallets")
        .sort_values([COL_CARRIER, COL_XDOCK, "pallets"], ascending=[True, True, False])
        .reset_index(drop=True)
    )

    # Matriz XDOCK vs Tipo (visual)
    pivot_xdock_tipo = (
        active.groupby([COL_XDOCK, COL_TIPO], dropna=False)
        .size()
        .reset_index(name="pallets")
        .pivot_table(index=COL_XDOCK, columns=COL_TIPO, values="pallets", aggfunc="sum", fill_value=0)
        .reset_index()
    )

    # Pendientes (tipos sin m² y xdock sin capacidad)
    missing_types = sorted(set(active[COL_TIPO].dropna().unique()) - set(PALLET_M2_BY_TYPE.keys()))
    missing_xdock = sorted(set(active[COL_XDOCK].dropna().unique()) - set(capacity_map.keys()))
    pendientes = []
    for t in missing_types:
        pendientes.append({"tipo": "TIPO DE PALLET sin m²", "valor": t})
    for x in missing_xdock:
        pendientes.append({"tipo": "XDOCK sin capacidad", "valor": x})
    pendientes_df = pd.DataFrame(pendientes)

    return by_xdock, by_carrier_xdock, cxt, pivot_xdock_tipo, pendientes_df


def fig_occupancy_barh(by_xdock: pd.DataFrame):
    # Barra horizontal placosona: % ocupación por XDOCK
    df = by_xdock.copy()
    df = df[df["ocupacion_%"].notna()].copy()
    if df.empty:
        fig, ax = plt.subplots(figsize=(10, 3))
        ax.text(0.5, 0.5, "Sin datos suficientes (capacidad o m² faltantes).", ha="center", va="center")
        ax.axis("off")
        return fig

    df = df.sort_values("ocupacion_%", ascending=True).copy()
    labels = df[COL_XDOCK].tolist()
    values = df["ocupacion_%"].tolist()

    fig, ax = plt.subplots(figsize=(11, max(3.5, 0.5 * len(labels) + 1)))
    ax.barh(labels, values)
    ax.set_xlabel("% Ocupación")
    ax.set_ylabel("Crossdock")
    ax.set_title("Ocupación por Crossdock (%)")

    # Valores al final de cada barra
    for i, v in enumerate(values):
        ax.text(v + 0.5, i, f"{v:.1f}%", va="center")

    ax.set_xlim(0, max(100, max(values) * 1.15))
    plt.tight_layout()
    return fig


def fig_pallets_by_type(active: pd.DataFrame):
    # Resumen de pallets por tipo (total, ya filtrado)
    tmp = active.groupby(COL_TIPO).size().reset_index(name="pallets").sort_values("pallets", ascending=True)
    if tmp.empty:
        fig, ax = plt.subplots(figsize=(10, 3))
        ax.text(0.5, 0.5, "Sin pallets activos.", ha="center", va="center")
        ax.axis("off")
        return fig

    fig, ax = plt.subplots(figsize=(10, max(3.5, 0.4 * len(tmp) + 1)))
    ax.barh(tmp[COL_TIPO], tmp["pallets"])
    ax.set_xlabel("Pallets")
    ax.set_ylabel("Tipo")
    ax.set_title("Pallets por Tipo (según filtros)")
    plt.tight_layout()
    return fig


def fig_to_png_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def df_to_reportlab_table(df: pd.DataFrame, max_rows=30):
    """
    Convierte df a Table reportlab con estilo básico.
    Limita filas para que sea ejecutivo y no infinito.
    """
    show = df.copy()
    if len(show) > max_rows:
        show = show.head(max_rows).copy()

    data = [list(show.columns)] + show.values.tolist()
    t = Table(data, repeatRows=1)
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F2937")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 9),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
                ("FONTSIZE", (0, 1), (-1, -1), 8),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F7F7")]),
            ]
        )
    )
    return t


def make_pdf_report_bytes(
    title: str,
    filters_text: str,
    kpi_dict: dict,
    occupancy_png: bytes,
    type_png: bytes,
    by_xdock: pd.DataFrame,
    carrier_xdock_type: pd.DataFrame,
):
    """
    PDF ejecutivo en landscape.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(letter), leftMargin=0.5 * inch, rightMargin=0.5 * inch,
                            topMargin=0.5 * inch, bottomMargin=0.5 * inch)

    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    story.append(Paragraph(filters_text, styles["Normal"]))
    story.append(Spacer(1, 10))

    # KPIs
    kpi_lines = [
        f"<b>Pallets activas:</b> {kpi_dict.get('pallets', 'N/A')}",
        f"<b>m² ocupados:</b> {kpi_dict.get('m2', 'N/A')}",
        f"<b>% ocupación global:</b> {kpi_dict.get('pct', 'N/A')}",
        f"<b>Fecha:</b> {kpi_dict.get('fecha', 'N/A')}",
    ]
    story.append(Paragraph(" | ".join(kpi_lines), styles["Normal"]))
    story.append(Spacer(1, 14))

    # Gráficos
    story.append(Paragraph("<b>Ocupación por Crossdock</b>", styles["Heading2"]))
    story.append(Image(io.BytesIO(occupancy_png), width=10.5 * inch, height=3.6 * inch))
    story.append(Spacer(1, 12))

    story.append(Paragraph("<b>Pallets por Tipo</b>", styles["Heading2"]))
    story.append(Image(io.BytesIO(type_png), width=10.5 * inch, height=3.2 * inch))
    story.append(Spacer(1, 12))

    # Tabla: resumen por xdock (ejecutivo)
    story.append(Paragraph("<b>Resumen por Crossdock (pallets, m², capacidad, %)</b>", styles["Heading2"]))
    table_df = by_xdock.copy()
    # redondeos bonitos
    if "m2_ocupados" in table_df.columns:
        table_df["m2_ocupados"] = table_df["m2_ocupados"].round(2)
    if "ocupacion_%" in table_df.columns:
        table_df["ocupacion_%"] = table_df["ocupacion_%"].round(2)
    if "capacidad_m2" in table_df.columns:
        table_df["capacidad_m2"] = table_df["capacidad_m2"].round(2)
    story.append(df_to_reportlab_table(table_df, max_rows=50))
    story.append(Spacer(1, 14))

    # Tabla: pallets por carrier + crossdock + tipo (ejecutivo)
    story.append(Paragraph("<b>Pallets por Carrier + Crossdock + Tipo (según filtros)</b>", styles["Heading2"]))
    cxt = carrier_xdock_type.copy()

    # Para que se vea ejecutivo y no infinito:
    # dejamos top 200 renglones por si está enorme
    story.append(df_to_reportlab_table(cxt, max_rows=200))
    story.append(Spacer(1, 10))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# =========================
# UI (EXECUTIVE)
# =========================
st.markdown("## GASO Comunicaciones — Reporte Ejecutivo de Ocupación de Crossdock")

st.caption("Regla: 1 fila = 1 pallet. No cuenta si ESTATUS DE SALIDA = SALIDA. Encabezado en fila 5 (A5:AC5).")

# --- Upload Excel
file = st.file_uploader("Sube tu Excel (.xlsx)", type=["xlsx"])
if not file:
    st.stop()

# Leer Excel
try:
    xls = pd.ExcelFile(file)
    sheet = st.selectbox("Selecciona hoja", xls.sheet_names, index=0)
    df_raw = pd.read_excel(xls, sheet_name=sheet, header=4).dropna(how="all").copy()
except Exception as e:
    st.error(f"No pude leer el Excel: {e}")
    st.stop()

df_mapped = normalize_and_map_columns(df_raw)
missing = ensure_required_columns(df_mapped)
if missing:
    st.error(f"Faltan columnas requeridas: {missing}")
    st.write("Columnas detectadas:", list(df_mapped.columns))
    st.stop()

# --- Sidebar: filtros
st.sidebar.header("Filtros")
carriers_detectados = sorted(df_mapped[COL_CARRIER].astype(str).str.strip().unique(), key=lambda x: x.upper())
carrier_opts = ["TODOS"] + carriers_detectados
default = [c for c in DEFAULT_TARGET_CARRIERS if c in carriers_detectados] or ["TODOS"]
carriers_filter = st.sidebar.multiselect("Carrier", options=carrier_opts, default=default)

# --- Capacidades: más gráfico (tarjetas)
st.sidebar.header("Capacidades (m²)")
xdocks = sorted(df_mapped[COL_XDOCK].astype(str).str.strip().unique())

# estado persistente
if "capacity_map" not in st.session_state:
    st.session_state.capacity_map = {}

# inicializa con preset
for x in xdocks:
    if x not in st.session_state.capacity_map:
        st.session_state.capacity_map[x] = float(PRESET_XDOCK_CAPACITY_M2.get(x, 0.0))

# UI “tarjetas” en sidebar: agrupadas en columnas
with st.sidebar.expander("Editar capacidades por Crossdock", expanded=False):
    cols = st.columns(1)  # sidebar, una columna; visualmente se siente como tarjetas
    for x in xdocks:
        st.session_state.capacity_map[x] = st.number_input(
            label=f"{x}",
            min_value=0.0,
            step=10.0,
            value=float(st.session_state.capacity_map.get(x, 0.0)),
            key=f"cap_{x}",
        )

capacity_map = dict(st.session_state.capacity_map)

# --- Construye activos y métricas (APLICA FILTROS)
try:
    active = build_active(df_raw, carriers_filter)
except Exception as e:
    st.error(f"Error procesando: {e}")
    st.stop()

by_xdock, by_carrier_xdock, cxt, pivot_xdock_tipo, pendientes_df = summarize(active, capacity_map)

# =========================
# BLOQUE 1: MATRIZ (XDOCK vs TIPO)
# =========================
st.markdown("### Matriz de Pallets por Crossdock y Tipo (activos)")
st.dataframe(pivot_xdock_tipo, use_container_width=True, hide_index=True)

if not pendientes_df.empty:
    st.warning("Pendientes detectados (pueden afectar m² o %):")
    st.dataframe(pendientes_df, use_container_width=True, hide_index=True)

# =========================
# BLOQUE 2: GRAFICO BARH % OCUPACIÓN
# =========================
st.markdown("### Ocupación por Crossdock (m² y %)")
fig1 = fig_occupancy_barh(by_xdock)
st.pyplot(fig1, use_container_width=True)

# =========================
# BLOQUE 3: GRAFICO PALLETS POR TIPO
# =========================
st.markdown("### Pallets por Tipo (según filtros)")
fig2 = fig_pallets_by_type(active)
st.pyplot(fig2, use_container_width=True)

# =========================
# KPI
# =========================
total_pallets = int(active["PALLETS_FILA"].sum())
total_m2 = float(active["M2_OCUPADOS_FILA"].sum(skipna=True))

cap_sum = by_xdock["capacidad_m2"].sum(skipna=True)
occ_sum = by_xdock["m2_ocupados"].sum(skipna=True)
pct_global = (occ_sum / cap_sum * 100.0) if cap_sum and cap_sum > 0 else float("nan")

k1, k2, k3 = st.columns(3)
k1.metric("Pallets activas", f"{total_pallets}")
k2.metric("m² ocupados", f"{total_m2:,.2f}")
k3.metric("% ocupación global", f"{pct_global:,.2f}%" if pd.notna(pct_global) else "N/A")

# =========================
# PDF EJECUTIVO (APLICA FILTROS)
# =========================
st.markdown("### Reporte Ejecutivo (PDF)")
filters_text = (
    f"Filtros aplicados — Carrier: "
    f"{', '.join(carriers_filter) if carriers_filter else 'TODOS'}"
)

title_pdf = "GASO Comunicaciones — Reporte Ejecutivo de Ocupación de Crossdock"
now_txt = datetime.now().strftime("%Y-%m-%d %H:%M")

kpi_dict = {
    "pallets": f"{total_pallets}",
    "m2": f"{total_m2:,.2f}",
    "pct": f"{pct_global:,.2f}%" if pd.notna(pct_global) else "N/A",
    "fecha": now_txt,
}

# recrea figs para export (porque st.pyplot ya pudo cerrarlas)
occ_png = fig_to_png_bytes(fig_occupancy_barh(by_xdock))
type_png = fig_to_png_bytes(fig_pallets_by_type(active))

pdf_bytes = make_pdf_report_bytes(
    title=title_pdf,
    filters_text=filters_text,
    kpi_dict=kpi_dict,
    occupancy_png=occ_png,
    type_png=type_png,
    by_xdock=by_xdock,
    carrier_xdock_type=cxt,  # esto es "cuantas palets de cada una hay por carrier y crossdock" (por tipo)
)

st.download_button(
    "⬇️ Descargar PDF Ejecutivo",
    data=pdf_bytes,
    file_name="GASO_reporte_ejecutivo_crossdock.pdf",
    mime="application/pdf",
)
