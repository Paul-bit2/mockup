import io
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
# CONFIG: capacidades m² por XDOCK (precargadas)
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

# Nota: tu columna S se llama "ESTATUS DE SALIDA"
REQUIRED_COLS = [
    "CARRIER",
    "XDOCK",
    "NO. DE PALLET",
    "TIPO DE PALLET",
    "ESTATUS DE SALIDA",
]

DEFAULT_TARGET_CARRIERS = ["TELCEL", "AT&T"]


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def ensure_required_columns(df: pd.DataFrame) -> list:
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    return missing


def is_in_inventory(row: pd.Series) -> bool:
    """
    Regla del usuario:
    - Si columna S (ESTATUS DE SALIDA) == "SALIDA" => ya no está en crossdock => NO contar
    - Cualquier otro caso => SÍ contar
    """
    estatus = row.get("ESTATUS DE SALIDA")
    estatus_txt = "" if pd.isna(estatus) else str(estatus).strip().upper()
    return estatus_txt != "SALIDA"


def safe_int(x, default=1) -> int:
    """
    NO. DE PALLET a entero.
    Si viene vacío o raro, asumimos 1 (porque tú cuentas por fila).
    """
    try:
        if pd.isna(x):
            return default
        s = str(x).strip()
        if s == "":
            return default
        # soporta cosas como "1 de 1" -> agarra el primer número
        # si viene "2" o "2.0" también sirve
        first_token = s.split()[0]
        v = int(float(first_token))
        return v
    except Exception:
        return default


def compute_row_m2(row: pd.Series) -> float:
    """
    m² por fila = (NO. DE PALLET) * (m² por tipo de pallet)
    En tu caso normalmente será 1 por fila, pero respetamos si trae un número distinto.
    """
    pallet_type = str(row.get("TIPO DE PALLET", "")).strip().upper()
    pallet_type = " ".join(pallet_type.split())  # normaliza espacios

    pallets = safe_int(row.get("NO. DE PALLET"), default=1)
    if pallets <= 0:
        pallets = 1

    if pallet_type not in PALLET_M2_BY_TYPE:
        return float("nan")

    return pallets * float(PALLET_M2_BY_TYPE[pallet_type])


def build_report(df_in: pd.DataFrame, capacity_map: dict, carriers_filter: list | None):
    df = normalize_columns(df_in)

    missing = ensure_required_columns(df)
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}")

    # Normaliza campos clave
    df["CARRIER"] = df["CARRIER"].astype(str).str.strip()
    df["XDOCK"] = df["XDOCK"].astype(str).str.strip()
    df["TIPO DE PALLET"] = (
        df["TIPO DE PALLET"]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )

    # Solo activos (según ESTATUS DE SALIDA != "SALIDA")
    active = df[df.apply(is_in_inventory, axis=1)].copy()

    # Filtra carriers si aplica
    if carriers_filter and len(carriers_filter) > 0 and "TODOS" not in carriers_filter:
        active["_CARRIER_UP"] = active["CARRIER"].str.upper()
        wanted = {c.upper() for c in carriers_filter}
        active = active[active["_CARRIER_UP"].isin(wanted)].copy()
        active.drop(columns=["_CARRIER_UP"], inplace=True)

    # Conteo por fila:
    # - Si NO. DE PALLET viene vacío o raro -> 1
    active["PALLETS_FILA"] = active["NO. DE PALLET"].apply(lambda x: safe_int(x, default=1))
    active.loc[active["PALLETS_FILA"] <= 0, "PALLETS_FILA"] = 1

    # m² y capacidad
    active["M2_OCUPADOS_FILA"] = active.apply(compute_row_m2, axis=1)
    active["CAPACIDAD_M2_XDOCK"] = active["XDOCK"].map(capacity_map)

    # Resumen por Carrier + XDOCK
    resumen = (
        active.groupby(["CARRIER", "XDOCK"], dropna=False)
        .agg(
            filas_contadas=("PALLETS_FILA", "count"),
            pallets_en_inventario=("PALLETS_FILA", "sum"),
            m2_ocupados=("M2_OCUPADOS_FILA", "sum"),
            capacidad_m2=("CAPACIDAD_M2_XDOCK", "first"),
        )
        .reset_index()
        .sort_values(["CARRIER", "XDOCK"])
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
    missing_types = sorted(set(active["TIPO DE PALLET"].dropna().unique()) - set(PALLET_M2_BY_TYPE.keys()))
    missing_xdock = sorted(set(active["XDOCK"].dropna().unique()) - set(capacity_map.keys()))

    pendientes_rows = []
    for t in missing_types:
        pendientes_rows.append({"tipo": "TIPO DE PALLET sin m²", "valor": t, "accion": "Agregar al catálogo de m²"})
    for x in missing_xdock:
        pendientes_rows.append({"tipo": "XDOCK sin capacidad", "valor": x, "accion": "Capturar capacidad m²"})

    pendientes = pd.DataFrame(pendientes_rows)

    # Detalle (inventario activo)
    front = [
        "CARRIER",
        "XDOCK",
        "TIPO DE PALLET",
        "NO. DE PALLET",
        "PALLETS_FILA",
        "M2_OCUPADOS_FILA",
        "CAPACIDAD_M2_XDOCK",
        "ESTATUS DE SALIDA",
    ]
    front = [c for c in front if c in active.columns]
    rest = [c for c in active.columns if c not in front]
    detalle = active[front + rest].copy()

    return resumen, detalle, pendientes


def to_excel_bytes(resumen, detalle, pendientes):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        detalle.to_excel(writer, sheet_name="Detalle_Activos", index=False)
        pendientes.to_excel(writer, sheet_name="Pendientes_Config", index=False)
    output.seek(0)
    return output.getvalue()


# =========================
# UI
# =========================
st.title("📦 Reporte de ocupación por Crossdock (m² y %)")

with st.expander("Reglas de conteo (importante)", expanded=False):
    st.markdown(
        "- **Encabezados en fila 5 (A5:AC5)** y datos desde **fila 6**.\n"
        "- **Cuenta por fila**, pero **NO cuenta** si **ESTATUS DE SALIDA = SALIDA**.\n"
        "- m² = **NO. DE PALLET × m² según TIPO DE PALLET**.\n"
        "- Si un tipo de pallet no existe en el catálogo de m², se marca como pendiente."
    )

file = st.file_uploader("Sube tu Excel (.xlsx)", type=["xlsx"])

if not file:
    st.stop()

# Leer Excel y hoja (header fila 5 => header=4)
try:
    xls = pd.ExcelFile(file)
    sheet = st.selectbox("Selecciona hoja", xls.sheet_names, index=0)
    df = pd.read_excel(xls, sheet_name=sheet, header=4)
    df = df.dropna(how="all").copy()
except Exception as e:
    st.error(f"No pude leer el Excel: {e}")
    st.stop()

df = normalize_columns(df)
missing = ensure_required_columns(df)
if missing:
    st.error(f"Faltan columnas requeridas: {missing}")
    st.write("Columnas detectadas:", list(df.columns))
    st.stop()

st.subheader("1) Capacidades por XDOCK (m²)")
xdocks = sorted(df["XDOCK"].astype(str).str.strip().unique())

# Tabla editable en sesión (precargada con preset)
if "cap_table" not in st.session_state:
    st.session_state.cap_table = pd.DataFrame(
        {
            "XDOCK": xdocks,
            "CAPACIDAD_M2": [float(PRESET_XDOCK_CAPACITY_M2.get(x, 0.0)) for x in xdocks],
        }
    )

# Si cambian los XDOCK (otro archivo), sincroniza conservando lo editado y usando preset si es nuevo
current = st.session_state.cap_table
if set(current["XDOCK"]) != set(xdocks):
    old_map = dict(zip(current["XDOCK"], current["CAPACIDAD_M2"]))
    st.session_state.cap_table = pd.DataFrame(
        {
            "XDOCK": xdocks,
            "CAPACIDAD_M2": [float(old_map.get(x, PRESET_XDOCK_CAPACITY_M2.get(x, 0.0))) for x in xdocks],
        }
    )

st.caption("Tip: puedes pegar valores desde Excel directo a esta tabla.")
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
carriers_detectados = sorted(df["CARRIER"].astype(str).str.strip().unique(), key=lambda x: x.upper())
opts = ["TODOS"] + carriers_detectados

default = [c for c in DEFAULT_TARGET_CARRIERS if c in carriers_detectados]
if not default:
    default = ["TODOS"]

carriers_filter = st.multiselect("Selecciona carriers", options=opts, default=default)

colA, colB = st.columns(2)
with colA:
    st.write("**Catálogo m² por tipo de pallet**")
    st.dataframe(
        pd.DataFrame({"TIPO DE PALLET": list(PALLET_M2_BY_TYPE.keys()), "M2": list(PALLET_M2_BY_TYPE.values())}),
        use_container_width=True,
        hide_index=True,
    )
with colB:
    st.write("**Vista rápida del archivo (primeras 30 filas)**")
    st.dataframe(df.head(30), use_container_width=True)

st.subheader("3) Generar reporte")
if st.button("📊 Calcular ocupación", type="primary"):
    try:
        resumen, detalle, pendientes = build_report(df, capacity_map, carriers_filter)
    except Exception as e:
        st.error(f"Error generando reporte: {e}")
        st.stop()

    # KPIs
    st.markdown("### KPIs rápidos")
    total_m2 = detalle["M2_OCUPADOS_FILA"].sum(skipna=True)
    total_pallets = detalle["PALLETS_FILA"].sum()

    k1, k2, k3 = st.columns(3)
    k1.metric("Tarimas (según conteo por fila)", f"{int(total_pallets)}")
    k2.metric("m² ocupados (total)", f"{total_m2:,.2f}")

    cap_sum = resumen["capacidad_m2"].sum(skipna=True)
    occ_sum = resumen["m2_ocupados"].sum(skipna=True)
    pct_global = (occ_sum / cap_sum * 100.0) if cap_sum and cap_sum > 0 else float("nan")
