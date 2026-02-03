import io
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Ocupación Crossdock", layout="wide")

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

# Nombres canónicos que usaremos dentro del programa
COL_CARRIER = "CARRIER"
COL_XDOCK = "XDOCK"
COL_TIPO = "TIPO DE PALLET"
COL_ESTATUS = "ESTATUS DE SALIDA"  # canónico


def canon(s: str) -> str:
    """
    Normaliza string para comparar nombres de columnas:
    - quita acentos
    - mayúsculas
    - colapsa espacios
    """
    if s is None:
        return ""
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.upper()
    s = " ".join(s.split())
    return s


def normalize_and_map_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    - Limpia nombres
    - Mapea variaciones a nombres canónicos (incluye ESTATUS DE SALIDA)
    """
    df = df.copy()
    original_cols = list(df.columns)
    clean_cols = [str(c).strip() for c in original_cols]
    df.columns = clean_cols

    # Mapa por "canon"
    # Llaves: posibles nombres en el Excel (normalizados)
    # Valores: nombre canónico interno
    alias_map = {
        # Carrier / XDOCK / Tipo
        canon("CARRIER"): COL_CARRIER,
        canon("XDOCK"): COL_XDOCK,
        canon("TIPO DE PALLET"): COL_TIPO,
        canon("TIPO PALLET"): COL_TIPO,

        # Estatus de salida (columna S)
        canon("ESTATUS DE SALIDA"): COL_ESTATUS,
        canon("ESTATUS SALIDA"): COL_ESTATUS,
        canon("STATUS DE SALIDA"): COL_ESTATUS,
        canon("STATUS SALIDA"): COL_ESTATUS,
        canon("ESTATUS SALIDA CLIENTE"): COL_ESTATUS,  # por si viene algo parecido
    }

    # Construye renombre real: col_actual -> col_canónica si coincide por canon
    rename_dict = {}
    for c in df.columns:
        key = canon(c)
        if key in alias_map:
            rename_dict[c] = alias_map[key]

    df = df.rename(columns=rename_dict)
    return df


def ensure_required_columns(df: pd.DataFrame) -> list:
    required = [COL_CARRIER, COL_XDOCK, COL_TIPO, COL_ESTATUS]
    return [c for c in required if c not in df.columns]


def is_in_inventory(row: pd.Series) -> bool:
    """
    Regla:
    - Si ESTATUS DE SALIDA == "SALIDA" => NO contar
    - Otro caso => SÍ contar
    """
    estatus = row.get(COL_ESTATUS)
    estatus_txt = "" if pd.isna(estatus) else str(estatus).strip().upper()
    return estatus_txt != "SALIDA"


def compute_row_m2(row: pd.Series) -> float:
    """
    1 fila = 1 pallet => m² depende SOLO del tipo.
    """
    pallet_type = str(row.get(COL_TIPO, "")).strip().upper()
    pallet_type = " ".join(pallet_type.split())

    if pallet_type not in PALLET_M2_BY_TYPE:
        return float("nan")
    return float(PALLET_M2_BY_TYPE[pallet_type])


def build_report(df_in: pd.DataFrame, capacity_map: dict, carriers_filter: list | None):
    df = normalize_and_map_columns(df_in)

    missing = ensure_required_columns(df)
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}. Columnas detectadas: {list(df.columns)}")

    # Normaliza campos clave
    df[COL_CARRIER] = df[COL_CARRIER].astype(str).str.strip()
    df[COL_XDOCK] = df[COL_XDOCK].astype(str).str.strip()
    df[COL_TIPO] = (
        df[COL_TIPO].astype(str).str.strip().str.upper().str.replace(r"\s+", " ", regex=True)
    )

    # Activos (sin SALIDA)
    active = df[df.apply(is_in_inventory, axis=1)].copy()

    # Filtra carriers
    if carriers_filter and len(carriers_filter) > 0 and "TODOS" not in carriers_filter:
        active["_CARRIER_UP"] = active[COL_CARRIER].str.upper()
        wanted = {c.upper() for c in carriers_filter}
        active = active[active["_CARRIER_UP"].isin(wanted)].copy()
        active.drop(columns=["_CARRIER_UP"], inplace=True)

    # 1 fila = 1 pallet
    active["PALLETS_FILA"] = 1

    # m² y capacidad
    active["M2_OCUPADOS_FILA"] = active.apply(compute_row_m2, axis=1)
    active["CAPACIDAD_M2_XDOCK"] = active[COL_XDOCK].map(capacity_map)

    resumen = (
        active.groupby([COL_CARRIER, COL_XDOCK], dropna=False)
        .agg(
            pallets_en_inventario=("PALLETS_FILA", "sum"),
            m2_ocupados=("M2_OCUPADOS_FILA", "sum"),
            capacidad_m2=("CAPACIDAD_M2_XDOCK", "first"),
        )
        .reset_index()
        .sort_values([COL_CARRIER, COL_XDOCK])
        .reset_index(drop=True)
    )

    def pct(row):
        cap = row["capacidad_m2"]
        occ = row["m2_ocupados"]
        if pd.isna(cap) or cap == 0 or pd.isna(occ):
            return float("nan")
        return (occ / cap) * 100.0

    resumen["ocupacion_%"] = resumen.apply(pct, axis=1)

    # Pendientes
    missing_types = sorted(set(active[COL_TIPO].dropna().unique()) - set(PALLET_M2_BY_TYPE.keys()))
    missing_xdock = sorted(set(active[COL_XDOCK].dropna().unique()) - set(capacity_map.keys()))

    pendientes_rows = []
    for t in missing_types:
        pendientes_rows.append({"tipo": "TIPO DE PALLET sin m²", "valor": t, "accion": "Agregar al catálogo de m²"})
    for x in missing_xdock:
        pendientes_rows.append({"tipo": "XDOCK sin capacidad", "valor": x, "accion": "Capturar capacidad m²"})
    pendientes = pd.DataFrame(pendientes_rows)

    # Conteos por XDOCK + Tipo
    conteo_xdock_tipo = (
        active.groupby([COL_XDOCK, COL_TIPO], dropna=False)
        .size()
        .reset_index(name="pallets_activas")
        .sort_values([COL_XDOCK, "pallets_activas"], ascending=[True, False])
        .reset_index(drop=True)
    )

    pivot_xdock_tipo = (
        conteo_xdock_tipo.pivot_table(
            index=COL_XDOCK,
            columns=COL_TIPO,
            values="pallets_activas",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    conteo_carrier_xdock_tipo = (
        active.groupby([COL_CARRIER, COL_XDOCK, COL_TIPO], dropna=False)
        .size()
        .reset_index(name="pallets_activas")
        .sort_values([COL_CARRIER, COL_XDOCK, "pallets_activas"], ascending=[True, True, False])
        .reset_index(drop=True)
    )

    # Detalle
    front = [
        COL_CARRIER,
        COL_XDOCK,
        COL_TIPO,
        "PALLETS_FILA",
        "M2_OCUPADOS_FILA",
        "CAPACIDAD_M2_XDOCK",
        COL_ESTATUS,
    ]
    front = [c for c in front if c in active.columns]
    rest = [c for c in active.columns if c not in front]
    detalle = active[front + rest].copy()

    return resumen, detalle, pendientes, conteo_xdock_tipo, pivot_xdock_tipo, conteo_carrier_xdock_tipo


def to_excel_bytes(resumen, detalle, pendientes, conteo_xdock_tipo, pivot_xdock_tipo, conteo_carrier_xdock_tipo):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        detalle.to_excel(writer, sheet_name="Detalle_Activos", index=False)
        pendientes.to_excel(writer, sheet_name="Pendientes_Config", index=False)
        conteo_xdock_tipo.to_excel(writer, sheet_name="Conteo_XDOCK_Tipo", index=False)
        pivot_xdock_tipo.to_excel(writer, sheet_name="Matriz_XDOCK_Tipo", index=False)
        conteo_carrier_xdock_tipo.to_excel(writer, sheet_name="Carrier_XDOCK_Tipo", index=False)
    output.seek(0)
    return output.getvalue()


# =========================
# UI
# =========================
st.title("📦 Reporte de ocupación por Crossdock (m² y %)")

with st.expander("Reglas de conteo (importante)", expanded=False):
    st.markdown(
        "- Encabezados en fila 5 (A5:AC5) y datos desde fila 6.\n"
        "- **1 fila = 1 pallet**.\n"
        "- **NO cuenta** si **ESTATUS DE SALIDA = SALIDA**.\n"
        "- m² por fila depende del **TIPO DE PALLET**.\n"
        "- La app detecta variantes del nombre de columna de salida (ej. `ESTATUS SALIDA`)."
    )

file = st.file_uploader("Sube tu Excel (.xlsx)", type=["xlsx"])
if not file:
    st.stop()

# Leer Excel (header fila 5 => header=4)
try:
    xls = pd.ExcelFile(file)
    sheet = st.selectbox("Selecciona hoja", xls.sheet_names, index=0)
    df_raw = pd.read_excel(xls, sheet_name=sheet, header=4).dropna(how="all").copy()
except Exception as e:
    st.error(f"No pude leer el Excel: {e}")
    st.stop()

df = normalize_and_map_columns(df_raw)
missing = ensure_required_columns(df)
if missing:
    st.error(f"Faltan columnas requeridas: {missing}")
    st.write("Columnas detectadas:", list(df.columns))
    st.stop()

st.subheader("1) Capacidades por XDOCK (m²)")
xdocks = sorted(df[COL_XDOCK].astype(str).str.strip().unique())

if "cap_table" not in st.session_state:
    st.session_state.cap_table = pd.DataFrame(
        {"XDOCK": xdocks, "CAPACIDAD_M2": [float(PRESET_XDOCK_CAPACITY_M2.get(x, 0.0)) for x in xdocks]}
    )

current = st.session_state.cap_table
if set(current["XDOCK"]) != set(xdocks):
    old_map = dict(zip(current["XDOCK"], current["CAPACIDAD_M2"]))
    st.session_state.cap_table = pd.DataFrame(
        {
            "XDOCK": xdocks,
            "CAPACIDAD_M2": [float(old_map.get(x, PRESET_XDOCK_CAPACITY_M2.get(x, 0.0))) for x in xdocks],
        }
    )

cap_edit = st.data_editor(
    st.session_state.cap_table,
    use_container_width=True,
    num_rows="fixed",
    column_config={
        "XDOCK": st.column_config.TextColumn(disabled=True),
        "CAPACIDAD_M2": st.column_config.NumberColumn(min_value=0.0, step=1.0),
    },
    hide_index=True,
)
st.session_state.cap_table = cap_edit
capacity_map = dict(zip(cap_edit["XDOCK"], cap_edit["CAPACIDAD_M2"]))

st.subheader("2) Filtros")
carriers_detectados = sorted(df[COL_CARRIER].astype(str).str.strip().unique(), key=lambda x: x.upper())
opts = ["TODOS"] + carriers_detectados
default = [c for c in DEFAULT_TARGET_CARRIERS if c in carriers_detectados] or ["TODOS"]
carriers_filter = st.multiselect("Selecciona carriers", options=opts, default=default)

st.subheader("3) Generar y ver conteos por tipo")
if st.button("📊 Calcular", type="primary"):
    try:
        resumen, detalle, pendientes, conteo_xdock_tipo, pivot_xdock_tipo, conteo_carrier_xdock_tipo = build_report(
            df, capacity_map, carriers_filter
        )
    except Exception as e:
        st.error(f"Error generando reporte: {e}")
        st.stop()

    st.markdown("### Conteo por XDOCK + TIPO (lista)")
    st.dataframe(conteo_xdock_tipo, use_container_width=True, hide_index=True)

    st.markdown("### Matriz XDOCK vs Tipo (más visual)")
    st.dataframe(pivot_xdock_tipo, use_container_width=True, hide_index=True)

    st.markdown("### Conteo por CARRIER + XDOCK + TIPO")
    st.dataframe(conteo_carrier_xdock_tipo, use_container_width=True, hide_index=True)

    st.markdown("### Resumen m² y % ocupación")
    st.dataframe(resumen, use_container_width=True, hide_index=True)

    st.markdown("### Pendientes (si falta mapear tipos o capacidad)")
    if pendientes.empty:
        st.success("✅ Todo ok.")
    else:
        st.warning("⚠️ Pendientes detectados.")
        st.dataframe(pendientes, use_container_width=True, hide_index=True)

    st.markdown("### Detalle (activos)")
    st.dataframe(detalle, use_container_width=True, hide_index=True)

    excel_bytes = to_excel_bytes(resumen, detalle, pendientes, conteo_xdock_tipo, pivot_xdock_tipo, conteo_carrier_xdock_tipo)
    st.download_button(
        "⬇️ Descargar reporte en Excel",
        data=excel_bytes,
        file_name="reporte_ocupacion_crossdock.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
