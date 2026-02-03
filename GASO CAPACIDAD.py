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

REQUIRED_COLS = [
    "CARRIER",
    "XDOCK",
    "TIPO DE PALLET",
    "ESTATUS DE SALIDA",
]

DEFAULT_TARGET_CARRIERS = ["TELCEL", "AT&T"]


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def ensure_required_columns(df: pd.DataFrame) -> list:
    return [c for c in REQUIRED_COLS if c not in df.columns]


def is_in_inventory(row: pd.Series) -> bool:
    """
    Regla:
    - Si ESTATUS DE SALIDA == "SALIDA" => NO contar
    - Otro caso => SÍ contar
    """
    estatus = row.get("ESTATUS DE SALIDA")
    estatus_txt = "" if pd.isna(estatus) else str(estatus).strip().upper()
    return estatus_txt != "SALIDA"


def compute_row_m2(row: pd.Series) -> float:
    """
    1 fila = 1 pallet => m² depende SOLO del TIPO DE PALLET.
    """
    pallet_type = str(row.get("TIPO DE PALLET", "")).strip().upper()
    pallet_type = " ".join(pallet_type.split())  # normaliza espacios

    if pallet_type not in PALLET_M2_BY_TYPE:
        return float("nan")

    return float(PALLET_M2_BY_TYPE[pallet_type])


def build_report(df_in: pd.DataFrame, capacity_map: dict, carriers_filter: list | None):
    df = normalize_columns(df_in)
    missing = ensure_required_columns(df)
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}")

    df["CARRIER"] = df["CARRIER"].astype(str).str.strip()
    df["XDOCK"] = df["XDOCK"].astype(str).str.strip()
    df["TIPO DE PALLET"] = (
        df["TIPO DE PALLET"]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )

    # Activos
    active = df[df.apply(is_in_inventory, axis=1)].copy()

    # Filtra carriers
    if carriers_filter and len(carriers_filter) > 0 and "TODOS" not in carriers_filter:
        active["_CARRIER_UP"] = active["CARRIER"].str.upper()
        wanted = {c.upper() for c in carriers_filter}
        active = active[active["_CARRIER_UP"].isin(wanted)].copy()
        active.drop(columns=["_CARRIER_UP"], inplace=True)

    # 1 fila = 1 pallet
    active["PALLETS_FILA"] = 1

    # m² y capacidad
    active["M2_OCUPADOS_FILA"] = active.apply(compute_row_m2, axis=1)
    active["CAPACIDAD_M2_XDOCK"] = active["XDOCK"].map(capacity_map)

    # Resumen por Carrier + XDOCK
    resumen = (
        active.groupby(["CARRIER", "XDOCK"], dropna=False)
        .agg(
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

    # Detalle
    front = [
        "CARRIER",
        "XDOCK",
        "TIPO DE PALLET",
        "PALLETS_FILA",
        "M2_OCUPADOS_FILA",
        "CAPACIDAD_M2_XDOCK",
        "ESTATUS DE SALIDA",
    ]
    front = [c for c in front if c in active.columns]
    rest = [c for c in active.columns if c not in front]
    detalle = active[front + rest].copy()

    return resumen, detalle, pendientes, active


def to_excel_bytes(resumen, detalle, pendientes, conteo_xdock_tipo, conteo_carrier_xdock_tipo):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        detalle.to_excel(writer, sheet_name="Detalle_Activos", index=False)
        pendientes.to_excel(writer, sheet_name="Pendientes_Config", index=False)
        conteo_xdock_tipo.to_excel(writer, sheet_name="Conteo_XDOCK_Tipo", index=False)
        conteo_carrier_xdock_tipo.to_excel(writer, sheet_name="Conteo_Carrier_XDOCK_Tipo", index=False)
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
        "- m² por fila depende del **TIPO DE PALLET**."
    )

file = st.file_uploader("Sube tu Excel (.xlsx)", type=["xlsx"])
if not file:
    st.stop()

# Leer Excel (header fila 5 => header=4)
try:
    xls = pd.ExcelFile(file)
    sheet = st.selectbox("Selecciona hoja", xls.sheet_names, index=0)
    df = pd.read_excel(xls, sheet_name=sheet, header=4).dropna(how="all").copy()
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

default = [c for c in DEFAULT_TARGET_CARRIERS if c in carriers_detectados] or ["TODOS"]
carriers_filter = st.multiselect("Selecciona carriers", options=opts, default=default)

# =========================
# 3) Conteos por XDOCK y TIPO (lo que pediste)
# =========================
st.subheader("3) Conteos de pallets ACTIVAS por Crossdock y Tipo")

tmp = df[df.apply(is_in_inventory, axis=1)].copy()
tmp["CARRIER"] = tmp["CARRIER"].astype(str).str.strip()
tmp["XDOCK"] = tmp["XDOCK"].astype(str).str.strip()
tmp["TIPO DE PALLET"] = (
    tmp["TIPO DE PALLET"].astype(str).str.strip().str.upper().str.replace(r"\s+", " ", regex=True)
)

# Conteo por XDOCK + Tipo de pallet
conteo_xdock_tipo = (
    tmp.groupby(["XDOCK", "TIPO DE PALLET"], dropna=False)
    .size()
    .reset_index(name="pallets_activas")
    .sort_values(["XDOCK", "pallets_activas"], ascending=[True, False])
    .reset_index(drop=True)
)

# Pivot: filas = XDOCK, columnas = tipo, valores = conteo
pivot_xdock_tipo = (
    conteo_xdock_tipo.pivot_table(
        index="XDOCK",
        columns="TIPO DE PALLET",
        values="pallets_activas",
        aggfunc="sum",
        fill_value=0,
    )
    .reset_index()
)

st.markdown("**A) Tabla por XDOCK y Tipo (lista)**")
st.dataframe(conteo_xdock_tipo, use_container_width=True, hide_index=True)

st.markdown("**B) Tabla tipo matriz (XDOCK vs Tipo)**")
st.dataframe(pivot_xdock_tipo, use_container_width=True, hide_index=True)

st.markdown("**C) Por Carrier + XDOCK + Tipo (lista)**")
conteo_carrier_xdock_tipo = (
    tmp.groupby(["CARRIER", "XDOCK", "TIPO DE PALLET"], dropna=False)
    .size()
    .reset_index(name="pallets_activas")
    .sort_values(["CARRIER", "XDOCK", "pallets_activas"], ascending=[True, True, False])
    .reset_index(drop=True)
)
st.dataframe(conteo_carrier_xdock_tipo, use_container_width=True, hide_index=True)

st.subheader("4) Generar reporte (m² y % ocupación)")
if st.button("📊 Calcular ocupación", type="primary"):
    try:
        resumen, detalle, pendientes, active = build_report(df, capacity_map, carriers_filter)
    except Exception as e:
        st.error(f"Error generando reporte: {e}")
        st.stop()

    st.markdown("### KPIs rápidos")
    total_pallets = int(detalle["PALLETS_FILA"].sum())
    total_m2 = float(detalle["M2_OCUPADOS_FILA"].sum(skipna=True))

    k1, k2, k3 = st.columns(3)
    k1.metric("Pallets activas (1 fila = 1)", f"{total_pallets}")
    k2.metric("m² ocupados (total)", f"{total_m2:,.2f}")

    cap_sum = resumen["capacidad_m2"].sum(skipna=True)
    occ_sum = resumen["m2_ocupados"].sum(skipna=True)
    pct_global = (occ_sum / cap_sum * 100.0) if cap_sum and cap_sum > 0 else float("nan")
    k3.metric("% ocupación global", f"{pct_global:,.2f}%" if pd.notna(pct_global) else "N/A")

    st.markdown("### Resumen por Carrier + XDOCK")
    st.dataframe(resumen, use_container_width=True, hide_index=True)

    st.markdown("### Detalle (solo inventario activo)")
    st.dataframe(detalle, use_container_width=True, hide_index=True)

    st.markdown("### Pendientes / alertas")
    if pendientes.empty:
        st.success("✅ Todo ok: no hay tipos de pallet ni XDOCKs pendientes.")
    else:
        st.warning("⚠️ Hay pendientes que debes completar para que el % sea correcto.")
        st.dataframe(pendientes, use_container_width=True, hide_index=True)

    excel_bytes = to_excel_bytes(resumen, detalle, pendientes, conteo_xdock_tipo, conteo_carrier_xdock_tipo)
    st.download_button(
        "⬇️ Descargar reporte en Excel",
        data=excel_bytes,
        file_name="reporte_ocupacion_crossdock.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
