import io
import re
import unicodedata
import streamlit.components.v1 as components
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
st.set_page_config(page_title="GASO Comunicaciones | Reportes", layout="wide")


# =========================
# UTIL: canon (global)
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


# ============================================================
# ===================== TAB 1: CROSSDOCK ======================
# ============================================================

# CONFIG: m² por tipo pallet
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

# CONFIG: capacidades m² por XDOCK
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
COL_SITE = "NOMBRE DE SITIO"


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
        canon("NOMBRE DE SITIO"): COL_SITE,
        canon("NOMBRE DEL SITIO"): COL_SITE,
        canon("NOMBRE SITIO"): COL_SITE,
        canon("SITIO"): COL_SITE,

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
    estatus = row.get(COL_ESTATUS)
    estatus_txt = "" if pd.isna(estatus) else str(estatus).strip().upper()
    return estatus_txt != "SALIDA"


def compute_row_m2(row: pd.Series) -> float:
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

    df[COL_CARRIER] = df[COL_CARRIER].astype(str).str.strip()
    df[COL_XDOCK] = df[COL_XDOCK].astype(str).str.strip()
    df[COL_TIPO] = df[COL_TIPO].astype(str).str.strip().str.upper().str.replace(r"\s+", " ", regex=True)

    active = df[df.apply(is_in_inventory, axis=1)].copy()

    if carriers_filter and len(carriers_filter) > 0 and "TODOS" not in carriers_filter:
        wanted = {c.upper() for c in carriers_filter}
        active = active[active[COL_CARRIER].str.upper().isin(wanted)].copy()

    active["PALLETS_FILA"] = 1
    active["M2_OCUPADOS_FILA"] = active.apply(compute_row_m2, axis=1)

    return active


def summarize(active: pd.DataFrame, capacity_map: dict):
    active = active.copy()
    active["CAPACIDAD_M2_XDOCK"] = active[COL_XDOCK].map(capacity_map)

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

    cxt = (
        active.groupby([COL_CARRIER, COL_XDOCK, COL_TIPO], dropna=False)
        .size()
        .reset_index(name="pallets")
        .sort_values([COL_CARRIER, COL_XDOCK, "pallets"], ascending=[True, True, False])
        .reset_index(drop=True)
    )

    pivot_xdock_tipo = (
        active.groupby([COL_XDOCK, COL_TIPO], dropna=False)
        .size()
        .reset_index(name="pallets")
        .pivot_table(index=COL_XDOCK, columns=COL_TIPO, values="pallets", aggfunc="sum", fill_value=0)
        .reset_index()
    )

    missing_types = sorted(set(active[COL_TIPO].dropna().unique()) - set(PALLET_M2_BY_TYPE.keys()))
    missing_xdock = sorted(set(active[COL_XDOCK].dropna().unique()) - set(capacity_map.keys()))
    pendientes = []
    for t in missing_types:
        pendientes.append({"tipo": "TIPO DE PALLET sin m²", "valor": t})
    for x in missing_xdock:
        pendientes.append({"tipo": "XDOCK sin capacidad", "valor": x})
    pendientes_df = pd.DataFrame(pendientes)

    return by_xdock, cxt, pivot_xdock_tipo, pendientes_df


def fig_occupancy_barh(by_xdock: pd.DataFrame):
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

    for i, v in enumerate(values):
        ax.text(v + 0.5, i, f"{v:.1f}%", va="center")

    ax.set_xlim(0, max(100, max(values) * 1.15))
    plt.tight_layout()
    return fig


def fig_pallets_by_type(active: pd.DataFrame):
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


def df_to_heatmap_pdf_table(df: pd.DataFrame, max_rows=30, exclude_cols=None):
    
    """
    Convierte DataFrame a tabla PDF con heatmap verde->amarillo->rojo
    (ignora primera columna si es texto tipo '# Económico')
    """
    show = df.copy()
    for col in show.columns:
        if pd.api.types.is_numeric_dtype(show[col]):
            show[col] = show[col].astype(float).round(2)
        if len(show) > max_rows:
            show = show.head(max_rows).copy()
        if exclude_cols is None:
            exclude_cols = ["Total general"]

    data = [list(show.columns)] + show.values.tolist()

    table = Table(data, repeatRows=1)
    table._argW = [1.2 * inch] + [0.55 * inch] * (len(show.columns) - 1)

        # ======= TABLA MÁS GRANDE (ajuste visual) =======
    styles = [
        ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F2937")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTSIZE", (0, 1), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
    ]

    # Columnas numéricas (ignoramos la primera, suele ser '# Económico')
    num_cols_idx = []
    for j, col in enumerate(show.columns):
        if j == 0:
            continue
        if pd.api.types.is_numeric_dtype(show[col]) and col != "Total general":
            num_cols_idx.append(j)

    # saca valores de MESES para escala (sin Total general)
    values = []
    for j in num_cols_idx:
        values.extend(pd.to_numeric(show.iloc[:, j], errors="coerce").dropna().tolist())

    # Si no hay números, retorna tabla con estilo básico
    if not values:
        table.setStyle(TableStyle(styles))
        return table

    vmin = float(min(values))
    vmax = float(max(values))
    span = (vmax - vmin) if vmax != vmin else 1.0

    def val_to_color(v, vmin_local, span_local):
        # normaliza 0-1
        x = (v - vmin_local) / span_local if span_local else 0.0
        x = max(0.0, min(1.0, x))

        # verde -> amarillo -> rojo (como RdYlGn_r)
        if x < 0.5:
            r = int(255 * (x * 2))
            g = 255
        else:
            r = 255
            g = int(255 * (1 - (x - 0.5) * 2))
        b = 0
        return colors.Color(r / 255, g / 255, b / 255)

    # ======= Heatmap MESES =======
    for row_i in range(1, len(show) + 1):
        for col_j in num_cols_idx:
            val = show.iloc[row_i - 1, col_j]
            val = pd.to_numeric(val, errors="coerce")
            if pd.notna(val):
                styles.append(
                    ("BACKGROUND", (col_j, row_i), (col_j, row_i), val_to_color(float(val), vmin, span))
                )

    # ======= Heatmap TOTAL GENERAL con escala propia =======
    if "Total general" in show.columns:
        tg_idx = list(show.columns).index("Total general")
        tg_vals = pd.to_numeric(show["Total general"], errors="coerce").dropna().tolist()
        tg_vals = [float(v) for v in tg_vals]

        if tg_vals:
            tg_vmin = float(min(tg_vals))
            tg_vmax = float(max(tg_vals))
            tg_span = (tg_vmax - tg_vmin) if tg_vmax != tg_vmin else 1.0

            for row_i in range(1, len(show) + 1):
                val = show.iloc[row_i - 1, tg_idx]
                val = pd.to_numeric(val, errors="coerce")
                if pd.notna(val):
                    styles.append(
                        ("BACKGROUND", (tg_idx, row_i), (tg_idx, row_i), val_to_color(float(val), tg_vmin, tg_span))
                    )

            styles.append(("FONTNAME", (tg_idx, 0), (tg_idx, -1), "Helvetica-Bold"))

    # ======= Anchos de columna (un poco más grande) =======
    # Primera columna más ancha, resto angosto pero legible
    table._argW = [1.2 * inch] + [0.52 * inch] * (len(show.columns) - 1)

    table.setStyle(TableStyle(styles))
    return table
def df_to_reportlab_table(df: pd.DataFrame, max_rows=30):
    """
    Convierte df a Table (ReportLab) con estilo ejecutivo.
    Limita filas para que no se haga infinito.
    """
    show = df.copy()

    # Redondeo general (2 decimales si es numérico)
    for c in show.columns:
        if pd.api.types.is_numeric_dtype(show[c]):
            show[c] = pd.to_numeric(show[c], errors="coerce").round(2)

    if len(show) > max_rows:
        show = show.head(max_rows).copy()

    data = [list(show.columns)] + show.values.tolist()

    t = Table(data, repeatRows=1)
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#111827")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 9),

                ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
                ("FONTSIZE", (0, 1), (-1, -1), 8),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F7F7")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
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
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch,
    )

    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    story.append(Paragraph(filters_text, styles["Normal"]))
    story.append(Spacer(1, 10))

    kpi_lines = [
        f"<b>Pallets activas:</b> {kpi_dict.get('pallets', 'N/A')}",
        f"<b>m² ocupados:</b> {kpi_dict.get('m2', 'N/A')}",
        f"<b>% ocupación global:</b> {kpi_dict.get('pct', 'N/A')}",
        f"<b>Fecha:</b> {kpi_dict.get('fecha', 'N/A')}",
    ]
    story.append(Paragraph(" | ".join(kpi_lines), styles["Normal"]))
    story.append(Spacer(1, 14))

    story.append(Paragraph("<b>Ocupación por Crossdock</b>", styles["Heading2"]))
    story.append(Image(io.BytesIO(occupancy_png), width=10.5 * inch, height=3.6 * inch))
    story.append(Spacer(1, 12))

    story.append(Paragraph("<b>Pallets por Tipo</b>", styles["Heading2"]))
    story.append(Image(io.BytesIO(type_png), width=10.5 * inch, height=3.2 * inch))
    story.append(Spacer(1, 12))

    story.append(Paragraph("<b>Resumen por Crossdock (pallets, m², capacidad, %)</b>", styles["Heading2"]))
    table_df = by_xdock.copy()
    if "m2_ocupados" in table_df.columns:
        table_df["m2_ocupados"] = table_df["m2_ocupados"].round(2)
    if "ocupacion_%" in table_df.columns:
        table_df["ocupacion_%"] = table_df["ocupacion_%"].round(2)
    if "capacidad_m2" in table_df.columns:
        table_df["capacidad_m2"] = table_df["capacidad_m2"].round(2)
    story.append(df_to_reportlab_table(table_df, max_rows=50))
    story.append(Spacer(1, 14))

    story.append(Paragraph("<b>Pallets por Carrier + Crossdock + Tipo (según filtros)</b>", styles["Heading2"]))
    story.append(df_to_reportlab_table(carrier_xdock_type.copy(), max_rows=200))
    story.append(Spacer(1, 10))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()
    
# ============================================================
# ==================== TAB 2: KILOMETRAJE =====================
# ============================================================

MONTHS_ES = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12
}

K_COL_MES = "MES"
K_COL_EQUIPO = "NOMBRE DEL EQUIPO"
K_COL_KM = "KILOMETRAJE"


def parse_mes_to_period(val) -> pd.Period:
    """
    Soporta:
    - 'ene-25', 'feb-25', 'ene-2025'
    - datetime/timestamp
    - serial Excel (número)
    """
    if pd.isna(val):
        return pd.NaT

    # ya viene como fecha
    if isinstance(val, (datetime, pd.Timestamp)):
        return pd.Period(val, freq="M")

    # serial Excel (int/float)
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        try:
            dt = pd.to_datetime(val, unit="D", origin="1899-12-30")
            return pd.Period(dt, freq="M")
        except Exception:
            pass

    s = str(val).strip().lower()
    s = s.replace(".", "").replace(" ", "").replace("_", "-").replace("/", "-")
    s = re.sub(r"[^a-z0-9\-]", "", s)

    m = re.match(r"^([a-z]{3})-?(\d{2}|\d{4})$", s)
    if not m:
        return pd.NaT

    mon, yr = m.group(1), m.group(2)

    if mon not in MONTHS_ES:
        return pd.NaT

    month = MONTHS_ES[mon]
    year = int(yr)
    if year < 100:
        year = 2000 + year

    return pd.Period(f"{year}-{month:02d}", freq="M")


def month_label(period: pd.Period) -> str:
    if pd.isna(period):
        return ""
    inv = {v: k for k, v in MONTHS_ES.items()}
    mon = inv.get(int(period.strftime("%m")), period.strftime("%m"))
    yy = period.strftime("%y")
    return f"{mon}-{yy}"


def normalize_km_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    alias = {
        canon("MES"): K_COL_MES,
        canon("NOMBRE DEL EQUIPO"): K_COL_EQUIPO,
        canon("EQUIPO"): K_COL_EQUIPO,
        canon("# ECONOMICO"): K_COL_EQUIPO,
        canon("ECONOMICO"): K_COL_EQUIPO,
        canon("KILOMETRAJE"): K_COL_KM,
        canon("KM"): K_COL_KM,
    }
    rename = {}
    for c in df.columns:
        k = canon(c)
        if k in alias:
            rename[c] = alias[k]
    return df.rename(columns=rename)


def build_km_dataset(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = normalize_km_columns(df_raw)
    missing = [c for c in [K_COL_MES, K_COL_EQUIPO, K_COL_KM] if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}. Columnas detectadas: {list(df.columns)}")

    df[K_COL_EQUIPO] = df[K_COL_EQUIPO].astype(str).str.strip()
    df[K_COL_MES] = df[K_COL_MES].apply(parse_mes_to_period)
    df = df[df[K_COL_MES].notna()].copy()

    def to_float(x):
        if pd.isna(x):
            return 0.0
        s = str(x).replace(",", "")
        try:
            return float(s)
        except Exception:
            return 0.0

    df[K_COL_KM] = df[K_COL_KM].apply(to_float)

    return df[[K_COL_EQUIPO, K_COL_MES, K_COL_KM]].copy()


def filter_period(df: pd.DataFrame, mode: str) -> pd.DataFrame:
    if df.empty:
        return df
    if mode == "Últimos 12 meses":
        max_p = df[K_COL_MES].max()
        start = (max_p - 11).asfreq("M")
        return df[df[K_COL_MES].between(start, max_p)].copy()
    return df.copy()


def monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(K_COL_MES).agg(
        suma_km=(K_COL_KM, "sum"),
        autos=(K_COL_KM, "size"),   # ← FILAS reales por mes
    ).reset_index()

    g["promedio_km_por_auto"] = g.apply(
        lambda r: (r["suma_km"] / r["autos"]) if r["autos"] else 0.0,
        axis=1
    )

    g["mes_label"] = g[K_COL_MES].apply(month_label)
    g = g.sort_values(K_COL_MES).reset_index(drop=True)

    return g[["mes_label", "suma_km", "autos", "promedio_km_por_auto"]]


def style_table_minmax(df: pd.DataFrame, col: str):
    if col not in df.columns or df.empty:
        return df.style
    vals = df[col]
    try:
        vmax = float(vals.max())
        vmin = float(vals.min())
    except Exception:
        return df.style

    def _color(v):
        try:
            v = float(v)
        except Exception:
            return ""
        if v == vmax:
            return "background-color: #ffcccc; font-weight: 700;"
        if v == vmin:
            return "background-color: #ccffcc; font-weight: 700;"
        return ""

    return df.style.applymap(_color, subset=[col])


def fig_barras_y_diferencia(month_df: pd.DataFrame):
    """
    Barras = suma_km
    Línea = autos
    Línea roja encima de barras (como pediste)
    """
    if month_df.empty:
        fig, ax = plt.subplots(figsize=(10, 3))
        ax.text(0.5, 0.5, "Sin datos.", ha="center", va="center")
        ax.axis("off")
        return fig

    df = month_df.copy()
    x = df["mes_label"].tolist()
    y = df["suma_km"].tolist()
    autos = df["autos"].tolist()

    fig, ax1 = plt.subplots(figsize=(12, 4))

    # barras (abajo)
    ax1.bar(x, y, zorder=1)
    ax1.set_axisbelow(True)
    ax1.set_ylabel("Suma de Kilometraje")
    ax1.set_title("Suma mensual de kilometraje y autos")

    # línea (arriba)
    ax2 = ax1.twinx()
    ax2.plot(
        x,
        autos,
        marker="o",
        color="red",     # cambia a "black" si la quieres negra
        linewidth=2.5,
        zorder=10
    )
    ax2.set_ylabel("Autos")

    # etiqueta de autos
    for i, val in enumerate(autos):
        ax2.text(i, val, f"{val}", ha="center", va="bottom")

    plt.tight_layout()
    return fig


def fig_promedio(month_df: pd.DataFrame):
    if month_df.empty:
        fig, ax = plt.subplots(figsize=(10, 3))
        ax.text(0.5, 0.5, "Sin datos.", ha="center", va="center")
        ax.axis("off")
        return fig
    fig, ax = plt.subplots(figsize=(12, 3.5))
    ax.plot(month_df["mes_label"], month_df["promedio_km_por_auto"], marker="o")
    ax.set_title("Promedio de km por auto (mensual)")
    ax.set_ylabel("Km promedio")
    plt.tight_layout()
    return fig


def top_acumulado(df: pd.DataFrame) -> pd.DataFrame:
    t = df.groupby(K_COL_EQUIPO)[K_COL_KM].sum().reset_index()
    t = t.rename(columns={K_COL_KM: "km_total"})
    t["servicios_requeridos"] = (t["km_total"] // 10000).astype(int)
    return t.sort_values("km_total", ascending=False).reset_index(drop=True)


def matrix_top_acumulado_formato(df: pd.DataFrame) -> pd.DataFrame:
    pv = df.pivot_table(index=K_COL_EQUIPO, columns=K_COL_MES, values=K_COL_KM, aggfunc="sum", fill_value=0.0)
    pv["Total general"] = pv.sum(axis=1)
    pv = pv.sort_values("Total general", ascending=False)
    pv.columns = [month_label(c) if isinstance(c, pd.Period) else str(c) for c in pv.columns]
    pv = pv.reset_index().rename(columns={K_COL_EQUIPO: "# Económico"})
    return pv


def top_mes(df: pd.DataFrame, mes_sel: pd.Period) -> pd.DataFrame:
    d = df[df[K_COL_MES] == mes_sel].copy()
    t = d.groupby(K_COL_EQUIPO)[K_COL_KM].sum().reset_index().rename(columns={K_COL_KM: "km_mes"})
    return t.sort_values("km_mes", ascending=False).reset_index(drop=True)


def gps_ceros(df: pd.DataFrame, mes_sel: pd.Period) -> pd.DataFrame:
    all_equip = df[K_COL_EQUIPO].dropna().unique().tolist()
    d = df[df[K_COL_MES] == mes_sel].groupby(K_COL_EQUIPO)[K_COL_KM].sum()
    rows = []
    for e in all_equip:
        km = float(d.get(e, 0.0))
        if km == 0.0:
            rows.append({K_COL_EQUIPO: e, "km_mes": 0.0})
    return pd.DataFrame(rows).sort_values(K_COL_EQUIPO).reset_index(drop=True)


def matrix_mes_equipo(df: pd.DataFrame) -> pd.DataFrame:
    pv = df.pivot_table(index=K_COL_EQUIPO, columns=K_COL_MES, values=K_COL_KM, aggfunc="sum", fill_value=0.0)
    pv.columns = [month_label(p) for p in pv.columns]
    return pv.reset_index().rename(columns={K_COL_EQUIPO: "Vehículo"})


def make_pdf_km_report_bytes(title, filtros_txt, kpis, fig1_png, fig2_png, tabla_mes, top_mat_df, top_total, top_mes_df, gps_df):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch
    )
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    story.append(Paragraph(filtros_txt, styles["Normal"]))
    story.append(Spacer(1, 10))

    story.append(Paragraph(
        " | ".join([
            f"<b>Autos:</b> {kpis.get('autos','N/A')}",
            f"<b>Km total:</b> {kpis.get('km_total','N/A')}",
            f"<b>Promedio mensual:</b> {kpis.get('prom_mensual','N/A')}",
            f"<b>Fecha:</b> {kpis.get('fecha','N/A')}",
        ]),
        styles["Normal"]
    ))
    story.append(Spacer(1, 12))

    story.append(Paragraph("<b>1) Suma mensual + autos (línea roja)</b>", styles["Heading2"]))
    story.append(Image(io.BytesIO(fig1_png), width=10.5 * inch, height=3.2 * inch))
    story.append(Spacer(1, 12))

    story.append(Paragraph("<b>2) Promedio km por auto</b>", styles["Heading2"]))
    story.append(Image(io.BytesIO(fig2_png), width=10.5 * inch, height=2.7 * inch))
    story.append(Spacer(1, 12))

    story.append(Paragraph("<b>3) Resumen mensual enero</b>", styles["Heading2"]))
    story.append(df_to_reportlab_table(tabla_mes, max_rows=24))
    story.append(Spacer(1, 12))

    story.append(Paragraph("<b>4) Top acumulado (matriz)</b>", styles["Heading2"]))
    story.append(df_to_heatmap_pdf_table(top_mat_df, max_rows=30, exclude_cols=[]))
    story.append(Spacer(1, 12))


    story.append(Paragraph("<b>4) Top acumulado + servicios que debería de tener cada auto</b>", styles["Heading2"]))
    story.append(df_to_reportlab_table(top_total, max_rows=25))
    story.append(Spacer(1, 12))

    story.append(Paragraph("<b>5) Top del mes</b>", styles["Heading2"]))
    story.append(df_to_reportlab_table(top_mes_df, max_rows=25))
    story.append(Spacer(1, 12))

    story.append(Paragraph("<b>6) Vehículos a revisar por falta de lectura en GPS (km=0)</b>", styles["Heading2"]))
    story.append(df_to_reportlab_table(gps_df, max_rows=40))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


def to_excel_km_bytes(tabla_mes, matriz_equipo_mes, top_acum, top_mes_df, gps_df, datos_km_df):
    """
    Excel completo + semi-dashboard.
    - NO recorta tablas
    - Incluye datos crudos completos (datos_km_df)
    - Incluye hoja Dashboard con gráfica (barras + línea)
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.chart.label import DataLabelList

    wb = Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    center = Alignment(horizontal="center", vertical="center")

    def add_sheet(name, df, freeze="A2"):
        ws = wb.create_sheet(title=name)
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
            ws.append(row)
            if r_idx == 1:
                for c_idx in range(1, len(row) + 1):
                    cell = ws.cell(row=1, column=c_idx)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = center

        ws.freeze_panes = freeze
        ws.auto_filter.ref = ws.dimensions

        for col_cells in ws.columns:
            max_len = 10
            col = col_cells[0].column_letter
            for c in col_cells:
                if c.value is None:
                    continue
                max_len = max(max_len, len(str(c.value)))
            ws.column_dimensions[col].width = min(60, max_len + 2)
        return ws

    # Datos crudos (TODO)
    add_sheet("Datos_Kilometraje", datos_km_df.copy(), freeze="A2")

    # Resumen mensual
    ws_m = add_sheet("Resumen_Mensual", tabla_mes.copy(), freeze="A2")
    ws_m.conditional_formatting.add(
        "B2:B1048576",
        ColorScaleRule(
            start_type="min", start_color="63BE7B",
            mid_type="percentile", mid_value=50, mid_color="FFEB84",
            end_type="max", end_color="F8696B"
        )
    )

    # Resto de hojas completas
    add_sheet("Top_Acumulado", top_acum.copy(), freeze="A2")
    add_sheet("Top_Mes", top_mes_df.copy(), freeze="A2")
    add_sheet("GPS_Revision", gps_df.copy(), freeze="A2")
    add_sheet("Matriz_Equipo_Mes", matriz_equipo_mes.copy(), freeze="B2")

    # Dashboard (semi)
    ws_dash = wb.create_sheet("Dashboard")
    ws_dash["A1"] = "GASO Comunicaciones — Dashboard de Kilometraje"
    ws_dash["A1"].font = Font(bold=True, size=16)

    ws_dash["A3"] = "Resumen mensual"
    ws_dash["A3"].font = Font(bold=True, size=12)

    start_row = 5
    for r_idx, row in enumerate(dataframe_to_rows(tabla_mes, index=False, header=True), start=start_row):
        ws_dash.append(row)
        if r_idx == start_row:
            for c_idx in range(1, len(row) + 1):
                cell = ws_dash.cell(row=r_idx, column=c_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center

    for col in ["A", "B", "C", "D"]:
        ws_dash.column_dimensions[col].width = 22

    last_row = start_row + len(tabla_mes)

    # barras: col B (Suma)
    bar = BarChart()
    bar.type = "col"
    bar.title = "Suma mensual de kilometraje"
    bar.y_axis.title = "Kilometraje"
    bar.dataLabels = DataLabelList()
    bar.dataLabels.showVal = False

    data_suma = Reference(ws_dash, min_col=2, min_row=start_row, max_row=last_row)  # incluye header
    cats = Reference(ws_dash, min_col=1, min_row=start_row + 1, max_row=last_row)
    bar.add_data(data_suma, titles_from_data=True)
    bar.set_categories(cats)

    # línea: col C (Autos)
    line = LineChart()
    line.y_axis.axId = 200
    line.y_axis.title = "Autos"
    line.y_axis.majorGridlines = None

    data_autos = Reference(ws_dash, min_col=3, min_row=start_row, max_row=last_row)
    line.add_data(data_autos, titles_from_data=True)
    line.set_categories(cats)

    bar += line
    ws_dash.add_chart(bar, "F5")
    ws_dash.freeze_panes = "A5"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

# =========================
# UI PRINCIPAL
# =========================
st.markdown("## GASO Comunicaciones — Dashboard Operativo")

tab_crossdock, tab_km, tab_sitio = st.tabs(["Crossdock", "Reporte kilometraje", "Pallets por Sitio"])


# -------------------------
# TAB CROSSDOCK
# -------------------------
with tab_crossdock:
    st.markdown("## GASO Comunicaciones — Reporte Ejecutivo de Ocupación de Crossdock")
    st.caption("Regla: 1 fila = 1 pallet. No cuenta si ESTATUS DE SALIDA = SALIDA. Encabezado en fila 5 (A5:AC5).")

    file_crossdock = st.file_uploader("Sube tu Excel de Crossdock (.xlsx)", type=["xlsx"], key="crossdock_excel")

    if not file_crossdock:
        st.info("👆 Sube un Excel para ver el reporte de Crossdock.")
    else:
        try:
            xls = pd.ExcelFile(file_crossdock)
            sheet = st.selectbox("Selecciona hoja", xls.sheet_names, index=0, key="crossdock_sheet")
            df_raw = pd.read_excel(xls, sheet_name=sheet, header=4).dropna(how="all").copy()
        except Exception as e:
            st.error(f"No pude leer el Excel: {e}")
            df_raw = None

        if df_raw is not None:
            df_mapped = normalize_and_map_columns(df_raw)
            missing = ensure_required_columns(df_mapped)
            if missing:
                st.error(f"Faltan columnas requeridas: {missing}")
                st.write("Columnas detectadas:", list(df_mapped.columns))
            else:
                st.sidebar.header("Crossdock — Filtros")
                carriers_detectados = sorted(df_mapped[COL_CARRIER].astype(str).str.strip().unique(), key=lambda x: x.upper())
                carrier_opts = ["TODOS"] + carriers_detectados
                default = [c for c in DEFAULT_TARGET_CARRIERS if c in carriers_detectados] or ["TODOS"]
                carriers_filter = st.sidebar.multiselect(
                    "Carrier (Crossdock)",
                    options=carrier_opts,
                    default=default,
                    key="crossdock_carrier"
                )

                st.sidebar.header("Crossdock — Capacidades (m²)")
                xdocks = sorted(df_mapped[COL_XDOCK].astype(str).str.strip().unique())

                if "capacity_map" not in st.session_state:
                    st.session_state.capacity_map = {}
                for x in xdocks:
                    if x not in st.session_state.capacity_map:
                        st.session_state.capacity_map[x] = float(PRESET_XDOCK_CAPACITY_M2.get(x, 0.0))

                with st.sidebar.expander("Editar capacidades por Crossdock", expanded=False):
                    for x in xdocks:
                        st.session_state.capacity_map[x] = st.number_input(
                            label=f"{x}",
                            min_value=0.0,
                            step=10.0,
                            value=float(st.session_state.capacity_map.get(x, 0.0)),
                            key=f"cap_{x}",
                        )

                capacity_map = dict(st.session_state.capacity_map)

                try:
                    active = build_active(df_raw, carriers_filter)
                except Exception as e:
                    st.error(f"Error procesando: {e}")
                    active = None

                if active is not None:
                    by_xdock, cxt, pivot_xdock_tipo, pendientes_df = summarize(active, capacity_map)

                    st.markdown("### Matriz de Pallets por Crossdock y Tipo (activos)")
                    st.dataframe(pivot_xdock_tipo, use_container_width=True, hide_index=True)

                    if not pendientes_df.empty:
                        st.warning("Pendientes detectados (pueden afectar m² o %):")
                        st.dataframe(pendientes_df, use_container_width=True, hide_index=True)

                    st.markdown("### Ocupación por Crossdock (m² y %)")
                    st.pyplot(fig_occupancy_barh(by_xdock), use_container_width=True)

                    st.markdown("### Pallets por Tipo (según filtros)")
                    st.pyplot(fig_pallets_by_type(active), use_container_width=True)

                    total_pallets = int(active["PALLETS_FILA"].sum())
                    total_m2 = float(active["M2_OCUPADOS_FILA"].sum(skipna=True))

                    cap_sum = by_xdock["capacidad_m2"].sum(skipna=True)
                    occ_sum = by_xdock["m2_ocupados"].sum(skipna=True)
                    pct_global = (occ_sum / cap_sum * 100.0) if cap_sum and cap_sum > 0 else float("nan")

                    k1, k2, k3 = st.columns(3)
                    k1.metric("Pallets activas", f"{total_pallets}")
                    k2.metric("m² ocupados", f"{total_m2:,.2f}")
                    k3.metric("% ocupación global", f"{pct_global:,.2f}%" if pd.notna(pct_global) else "N/A")

                    st.markdown("### Reporte Ejecutivo (PDF)")
                    filters_text = f"Filtros aplicados — Carrier: {', '.join(carriers_filter) if carriers_filter else 'TODOS'}"
                    title_pdf = "GASO Comunicaciones — Reporte Ejecutivo de Ocupación de Crossdock"
                    now_txt = datetime.now().strftime("%Y-%m-%d %H:%M")

                    kpi_dict = {
                        "pallets": f"{total_pallets}",
                        "m2": f"{total_m2:,.2f}",
                        "pct": f"{pct_global:,.2f}%" if pd.notna(pct_global) else "N/A",
                        "fecha": now_txt,
                    }

                    occ_png = fig_to_png_bytes(fig_occupancy_barh(by_xdock))
                    type_png = fig_to_png_bytes(fig_pallets_by_type(active))

                    pdf_bytes = make_pdf_report_bytes(
                        title=title_pdf,
                        filters_text=filters_text,
                        kpi_dict=kpi_dict,
                        occupancy_png=occ_png,
                        type_png=type_png,
                        by_xdock=by_xdock,
                        carrier_xdock_type=cxt,
                    )

                    st.download_button(
                        "⬇️ Descargar PDF Ejecutivo (Crossdock)",
                        data=pdf_bytes,
                        file_name="GASO_reporte_ejecutivo_crossdock.pdf",
                        mime="application/pdf",
                    )

# -------------------------
# TAB KILOMETRAJE
# -------------------------
with tab_km:
    st.markdown("## GASO Comunicaciones — Reporte Ejecutivo de Kilometraje")
    st.caption("Formato: Item | Mes | Nombre del equipo | Kilometraje (A1 encabezado, datos desde A2).")

    km_file = st.file_uploader("Sube tu Excel de kilometraje (.xlsx)", type=["xlsx"], key="km_uploader")

    if not km_file:
        st.info("👆 Sube un Excel para ver el reporte de Kilometraje.")
    else:
        try:
            xls = pd.ExcelFile(km_file)
            sheet = st.selectbox("Selecciona hoja", xls.sheet_names, index=0, key="km_sheet")
            df_km_raw = pd.read_excel(xls, sheet_name=sheet, header=0).dropna(how="all").copy()
            df_km = build_km_dataset(df_km_raw)
        except Exception as e:
            st.error(f"No pude leer/procesar el Excel de kilometraje: {e}")
            df_km = None

        if df_km is not None:
            st.sidebar.header("Kilometraje — Filtros")
            mode_period = st.sidebar.radio(
                "Rango de análisis (Kilometraje)",
                ["Todo", "Últimos 12 meses"],
                index=0,
                key="km_range"
            )

            df_km_f = filter_period(df_km, mode_period)

            meses = sorted(df_km_f[K_COL_MES].unique())
            if not meses:
                st.error("No hay meses en el rango seleccionado. Revisa que la columna 'Mes' venga como ene-25, feb-25, etc.")
            else:
                mes_labels = [month_label(p) for p in meses]
                mes_sel_label = st.sidebar.selectbox(
                    "Mes para Top del mes y GPS",
                    mes_labels,
                    index=len(mes_labels) - 1,
                    key="km_month"
                )
                mes_sel = meses[mes_labels.index(mes_sel_label)]

                t_mes = monthly_summary(df_km_f)
                t_mes_display = t_mes.rename(columns={
                    "mes_label": "Mes",
                    "suma_km": "Suma de kilometraje",
                    "autos": "Autos",
                    "promedio_km_por_auto": "Promedio (km/auto)"
                })

                st.markdown("### 1) Resumen mensual")
                st.dataframe(style_table_minmax(t_mes_display, "Suma de kilometraje"), use_container_width=True, hide_index=True)

                st.markdown("### 2) Gráfica: barras (kilometraje) + línea (autos)")
                st.pyplot(fig_barras_y_diferencia(t_mes), use_container_width=True)

                st.markdown("### 3) TOP ACUMULADO")
                top_mat = matrix_top_acumulado_formato(df_km_f)
                num_cols = [c for c in top_mat.columns if c != "# Económico"]
                st.dataframe(
                    top_mat.style.background_gradient(cmap="RdYlGn_r", subset=num_cols),
                    use_container_width=True,
                    hide_index=True
                )

                st.markdown("### 4) Servicios requeridos (cada 10,000 km)")
                top_total = top_acumulado(df_km_f).rename(columns={
                    K_COL_EQUIPO: "Vehículo",
                    "km_total": "Km total",
                    "servicios_requeridos": "Servicios (floor)"
                })
                st.dataframe(top_total, use_container_width=True, hide_index=True)

                st.markdown(f"### 5) Top del mes seleccionado — {mes_sel_label}")
                top_m = top_mes(df_km_f, mes_sel).rename(columns={K_COL_EQUIPO: "Vehículo", "km_mes": "Km del mes"})
                st.dataframe(top_m.head(50), use_container_width=True, hide_index=True)

                st.markdown(f"### 6) Vehículos a revisar GPS — {mes_sel_label} (Km = 0)")
                gps_df = gps_ceros(df_km_f, mes_sel).rename(columns={K_COL_EQUIPO: "Vehículo", "km_mes": "Km del mes"})
                st.dataframe(gps_df, use_container_width=True, hide_index=True)

                total_autos = int(df_km_f[K_COL_EQUIPO].nunique())
                km_total = float(df_km_f[K_COL_KM].sum())
                prom_mensual = float(t_mes["suma_km"].mean()) if not t_mes.empty else 0.0

                k1, k2, k3 = st.columns(3)
                k1.metric("Autos en análisis", f"{total_autos}")
                k2.metric("Km total (rango)", f"{km_total:,.2f}")
                k3.metric("Promedio mensual (km)", f"{prom_mensual:,.2f}")

                st.markdown("### Descargas")
                matriz_eq_mes = matrix_mes_equipo(df_km_f)

                datos_km_df = df_km_f.copy()
                datos_km_df["MES_LABEL"] = datos_km_df["MES"].apply(month_label)
                datos_km_df = datos_km_df.rename(columns={"NOMBRE DEL EQUIPO": "Vehículo", "KILOMETRAJE": "Kilometraje"})
                datos_km_df = datos_km_df[["Vehículo", "MES_LABEL", "Kilometraje"]].sort_values(["MES_LABEL", "Vehículo"])

                excel_bytes = to_excel_km_bytes(
                    tabla_mes=t_mes_display,
                    matriz_equipo_mes=matriz_eq_mes,
                    top_acum=top_total,
                    top_mes_df=top_m,
                    gps_df=gps_df,
                    datos_km_df=datos_km_df
                )
                st.download_button(
                    "⬇️ Descargar Excel (Kilometraje completo + Dashboard)",
                    data=excel_bytes,
                    file_name=f"GASO_kilometraje_completo_{mode_period.replace(' ', '_')}_{mes_sel_label}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                now_txt = datetime.now().strftime("%Y-%m-%d %H:%M")
                pdf_bytes = make_pdf_km_report_bytes(
                    title="GASO Comunicaciones — Reporte Ejecutivo de Kilometraje",
                    filtros_txt=f"Rango: {mode_period} | Mes seleccionado: {mes_sel_label}",
                    kpis={"autos": f"{total_autos}", "km_total": f"{km_total:,.2f}", "prom_mensual": f"{prom_mensual:,.2f}", "fecha": now_txt},
                    fig1_png=fig_to_png_bytes(fig_barras_y_diferencia(t_mes)),
                    fig2_png=fig_to_png_bytes(fig_promedio(t_mes)),
                    tabla_mes=t_mes_display,
                    top_mat_df=top_mat,
                    top_total=top_total,
                    top_mes_df=top_m,
                    gps_df=gps_df
                )
                st.download_button(
                    "⬇️ Descargar PDF Ejecutivo (Kilometraje)",
                    data=pdf_bytes,
                    file_name=f"GASO_kilometraje_ejecutivo_{mode_period.replace(' ', '_')}_{mes_sel_label}.pdf",
                    mime="application/pdf"
                )
# -------------------------
# TAB PALLETS POR SITIO
# -------------------------
with tab_sitio:
    st.markdown("## Pallets por Sitio (tabla bonita para correo)")
    st.caption("Regla: 1 fila = 1 pallet. No cuenta si ESTATUS DE SALIDA = SALIDA. Encabezado en fila 5 (A5:AC5).")

    file_s = st.file_uploader("Sube tu Excel (.xlsx) para este reporte", type=["xlsx"], key="sitios_excel_tab")

    if not file_s:
        st.info("👆 Sube un Excel para ver el reporte por Sitio.")
    else:
        try:
            xls_s = pd.ExcelFile(file_s)
            sheet_s = st.selectbox("Selecciona hoja", xls_s.sheet_names, index=0, key="sitios_sheet")
            df_raw_s = pd.read_excel(xls_s, sheet_name=sheet_s, header=4).dropna(how="all").copy()
        except Exception as e:
            st.error(f"No pude leer el Excel: {e}")
            df_raw_s = None

        if df_raw_s is not None:
            df_map = normalize_and_map_columns(df_raw_s)
            needed = [COL_CARRIER, COL_XDOCK, COL_ESTATUS, COL_SITE]
            missing = [c for c in needed if c not in df_map.columns]
            if missing:
                st.error(f"Faltan columnas requeridas para este reporte: {missing}")
                st.write("Columnas detectadas:", list(df_map.columns))
            else:
                xdocks = sorted(df_map[COL_XDOCK].astype(str).str.strip().unique())
                xdock_sel = st.selectbox("Crossdock", options=["TODOS"] + xdocks, index=0, key="sitios_xdock")

                active_all = build_active(df_raw_s, ["TODOS"])
                if xdock_sel != "TODOS":
                    active_all = active_all[active_all[COL_XDOCK].astype(str).str.strip() == xdock_sel].copy()
                active_all[COL_SITE] = active_all[COL_SITE].astype(str).str.strip()

                sitios_lista = (
                    active_all.groupby(COL_SITE)
                    .size()
                    .reset_index(name="Pallets")
                    .sort_values("Pallets", ascending=False)[COL_SITE]
                    .astype(str)
                    .tolist()
                )

                excluir = st.multiselect("Quitar sitios (prueba/dummy) antes de generar", options=sitios_lista, default=[], key="sitios_excluir")

                if "sitios_manual_rows" not in st.session_state:
                    st.session_state.sitios_manual_rows = []

                with st.expander("Agregar sitios manuales (opcional)", expanded=False):
                    c1, c2, c3 = st.columns([2, 1, 1])
                    manual_site = c1.text_input("Nombre del sitio", key="sitio_manual_nombre")
                    manual_pallets = c2.number_input("Pallets", min_value=0, step=1, value=0, key="sitio_manual_pallets")
                    if c3.button("Agregar", key="sitio_manual_add"):
                        if str(manual_site).strip():
                            st.session_state.sitios_manual_rows.append({COL_SITE: str(manual_site).strip(), "Pallets": int(manual_pallets)})

                    if st.session_state.sitios_manual_rows:
                        st.dataframe(pd.DataFrame(st.session_state.sitios_manual_rows), hide_index=True, use_container_width=True)
                        if st.button("Limpiar manuales", key="sitio_manual_clear"):
                            st.session_state.sitios_manual_rows = []

                def build_pivot(carrier_filter: list[str]) -> pd.DataFrame:
                    active = build_active(df_raw_s, carrier_filter)
                    if xdock_sel != "TODOS":
                        active = active[active[COL_XDOCK].astype(str).str.strip() == xdock_sel].copy()

                    active[COL_SITE] = active[COL_SITE].astype(str).str.strip()
                    piv = (
                        active.groupby(COL_SITE, dropna=False)
                        .size()
                        .reset_index(name="Pallets")
                        .sort_values("Pallets", ascending=False)
                        .reset_index(drop=True)
                    )

                    if excluir:
                        piv = piv[~piv[COL_SITE].isin(excluir)].copy()

                    if st.session_state.sitios_manual_rows:
                        manual_df = pd.DataFrame(st.session_state.sitios_manual_rows)
                        piv = pd.concat([piv, manual_df], ignore_index=True)
                        piv = piv.groupby(COL_SITE, dropna=False)["Pallets"].sum().reset_index()
                        piv = piv.sort_values("Pallets", ascending=False).reset_index(drop=True)

                    total = int(piv["Pallets"].sum()) if not piv.empty else 0
                    piv = pd.concat([piv, pd.DataFrame([{COL_SITE: "TOTAL", "Pallets": total}])], ignore_index=True)
                    return piv

                st.markdown("### Generar tabla (copiar y pegar al correo)")
                b1, b2, b3 = st.columns(3)

                def render(title: str, df_table: pd.DataFrame):
                    html = df_to_email_html_table(df_table, title=title)
                    height = min(900, 140 + (len(df_table) * 28))
                    components.html(html, height=height, scrolling=True)

                with b1:
                    if st.button("📋 AMBOS", key="btn_sitios_ambos"):
                        piv = build_pivot(["TODOS"])
                        render(f"Pallets por Sitio — TODOS — XDOCK: {xdock_sel}", piv)

                with b2:
                    if st.button("📋 TELCEL", key="btn_sitios_telcel"):
                        piv = build_pivot(["TELCEL"])
                        render(f"Pallets por Sitio — TELCEL — XDOCK: {xdock_sel}", piv)

                with b3:
                    if st.button("📋 AT&T", key="btn_sitios_att"):
                        piv = build_pivot(["AT&T"])
                        render(f"Pallets por Sitio — AT&T — XDOCK: {xdock_sel}", piv)
