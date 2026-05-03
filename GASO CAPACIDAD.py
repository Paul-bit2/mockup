"""
GASO COMUNICACIONES – IN-OUT Report Processor
Cleans, enriches and transforms the IN-OUT Excel database and generates
an executive client report.

Persistent decision memory: manual XDOCK assignments are saved to
'gaso_decisions.json' so they auto-apply on future uploads.
"""

import io
import json
import os
import re
import datetime
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from unidecode import unidecode

# ─────────────────────────────────────────────────────────────────────────────
#  CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
DECISIONS_FILE = "gaso_decisions.json"   # persisted next to the script

GASO_BLUE   = "#1A3A6B"
GASO_LIGHT  = "#2E6DB4"
GASO_ACCENT = "#4A90D9"

CAPACIDADES = {
    "Gaso- Tijuana-E-NS":      677,
    "Gaso- La Paz-E-NS":       316,
    "Gaso- Hermosillo-E-NS":   400,
    "Gaso- Culiacán-E-NS":     600,
    "Gaso- Cd. Juarez-E-NS":   350,
    "Gaso- Chihuahua-E-NS":    400,
    "Gaso- Monterrey-E-NS":   1500,
    "Gaso- Guadalajara-E-NS": 1200,
    "Gaso- Querétaro-E-NS":    800,
    "Gaso- Torreón-E-NS":      250,
}

REGION_MAP = {
    "Gaso- La Paz-E-NS":       "REGIÓN LUIS",
    "Gaso- Culiacán-E-NS":     "REGIÓN LUIS",
    "Gaso- Guadalajara-E-NS":  "REGIÓN LUIS",
    "Gaso- Querétaro-E-NS":    "REGIÓN LUIS",
    "Gaso- Tijuana-E-NS":      "REGIÓN JORGE",
    "Gaso- Hermosillo-E-NS":   "REGIÓN JORGE",
    "Gaso- Cd. Juarez-E-NS":   "REGIÓN JORGE",
    "Gaso- Chihuahua-E-NS":    "REGIÓN JORGE",
    "Gaso- Monterrey-E-NS":    "REGIÓN JORGE",
    "Gaso- Torreón-E-NS":      "REGIÓN JORGE",
}

CIUDAD_MAP = {
    "Gaso- Tijuana-E-NS":      "Tijuana",
    "Gaso- La Paz-E-NS":       "La Paz",
    "Gaso- Hermosillo-E-NS":   "Hermosillo",
    "Gaso- Culiacán-E-NS":     "Culiacán",
    "Gaso- Cd. Juarez-E-NS":   "Cd. Juárez",
    "Gaso- Chihuahua-E-NS":    "Chihuahua",
    "Gaso- Monterrey-E-NS":    "Monterrey",
    "Gaso- Guadalajara-E-NS":  "Guadalajara",
    "Gaso- Querétaro-E-NS":    "Querétaro",
    "Gaso- Torreón-E-NS":      "Torreón",
}

M2_TIPO = {
    "EUROPALLET":        0.96,
    "ESTANDAR":          1.44,
    "SOBREDIMENSIONADA": 3.66,
}

FOLIO_XDOCK = [
    ("GTOR",  "Gaso- Torreón-E-NS"),
    ("GCCJS", "Gaso- Cd. Juarez-E-NS"),
    ("GASO",  "Gaso- Chihuahua-E-NS"),
    ("CUL",   "Gaso- Culiacán-E-NS"),
    ("GCH",   "Gaso- Hermosillo-E-NS"),
    ("BS",    "Gaso- La Paz-E-NS"),
    ("GCM",   "Gaso- Monterrey-E-NS"),
    ("GCQ",   "Gaso- Querétaro-E-NS"),
    ("GCTIJ", "Gaso- Tijuana-E-NS"),
]

ID_SITIO_PREFIXES = {
    "JAL": "Gaso- Guadalajara-E-NS",
    "SIN": "Gaso- Culiacán-E-NS",
    "NL":  "Gaso- Monterrey-E-NS",
    "NLE": "Gaso- Monterrey-E-NS",
}

KEYWORD_XDOCK = {
    "culiacan":    "Gaso- Culiacán-E-NS",
    "culiacán":    "Gaso- Culiacán-E-NS",
    "gdl":         "Gaso- Guadalajara-E-NS",
    "guadalajara": "Gaso- Guadalajara-E-NS",
    "mty":         "Gaso- Monterrey-E-NS",
    "monterrey":   "Gaso- Monterrey-E-NS",
    "qro":         "Gaso- Querétaro-E-NS",
    "queretaro":   "Gaso- Querétaro-E-NS",
    "querétaro":   "Gaso- Querétaro-E-NS",
    "hermosillo":  "Gaso- Hermosillo-E-NS",
    "juarez":      "Gaso- Cd. Juarez-E-NS",
    "juárez":      "Gaso- Cd. Juarez-E-NS",
    "chihuahua":   "Gaso- Chihuahua-E-NS",
    "torreon":     "Gaso- Torreón-E-NS",
    "torreón":     "Gaso- Torreón-E-NS",
    "tijuana":     "Gaso- Tijuana-E-NS",
    "la paz":      "Gaso- La Paz-E-NS",
    "lapaz":       "Gaso- La Paz-E-NS",
}

PALLET_NORM = {
    "antena gd":         "SOBREDIMENSIONADA",
    "antena md":         "ESTANDAR",
    "antena ch":         "EUROPALLET",
    "gabinete gd":       "ESTANDAR",
    "gabinete med":      "ESTANDAR",
    "galvanizado":       "SOBREDIMENSIONADA",
    "soporte":           "SOBREDIMENSIONADA",
    "estandar":          "ESTANDAR",
    "europallet":        "EUROPALLET",
    "sobredimensionado": "SOBREDIMENSIONADA",
    "sobredimensionada": "SOBREDIMENSIONADA",
}

XDOCK_ALIASES = {
    "gaso- torreon-e-ns":     "Gaso- Torreón-E-NS",
    "gaso- torreón- e-ns":    "Gaso- Torreón-E-NS",
    "gaso- torreon- e-ns":    "Gaso- Torreón-E-NS",
    "gaso- tijuana-e-ns":     "Gaso- Tijuana-E-NS",
    "gaso- la paz-e-ns":      "Gaso- La Paz-E-NS",
    "gaso- hermosillo-e-ns":  "Gaso- Hermosillo-E-NS",
    "gaso- culiacan-e-ns":    "Gaso- Culiacán-E-NS",
    "gaso- culiacán-e-ns":    "Gaso- Culiacán-E-NS",
    "gaso- cd. juarez-e-ns":  "Gaso- Cd. Juarez-E-NS",
    "gaso- chihuahua-e-ns":   "Gaso- Chihuahua-E-NS",
    "gaso- monterrey-e-ns":   "Gaso- Monterrey-E-NS",
    "gaso- guadalajara-e-ns": "Gaso- Guadalajara-E-NS",
    "gaso- queretaro-e-ns":   "Gaso- Querétaro-E-NS",
    "gaso- querétaro-e-ns":   "Gaso- Querétaro-E-NS",
}

XDOCK_OPTIONS = ["— Selecciona —"] + sorted(CAPACIDADES.keys())

# ─────────────────────────────────────────────────────────────────────────────
#  PERSISTENT DECISIONS  (saved to gaso_decisions.json beside the script)
# ─────────────────────────────────────────────────────────────────────────────
def load_decisions() -> dict:
    """Load saved ID_SITIO → {action, xdock} decisions from disk."""
    if os.path.exists(DECISIONS_FILE):
        try:
            with open(DECISIONS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_decisions(decisions: dict):
    """Persist decisions to disk."""
    with open(DECISIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(decisions, f, ensure_ascii=False, indent=2)


def delete_decision(id_sitio: str):
    d = load_decisions()
    if id_sitio in d:
        del d[id_sitio]
        save_decisions(d)


# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def norm(s):
    if s is None:
        return ""
    return unidecode(str(s)).strip().lower()


def is_empty_xdock(val):
    n = norm(val)
    return not n or n.startswith("seleccion") or n in ("none", "nan")


def contains_ens(val):
    return "e-ns" in norm(val)


# ─────────────────────────────────────────────────────────────────────────────
#  LOAD
# ─────────────────────────────────────────────────────────────────────────────
def load_file(uploaded_file):
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb["IN-OUT"]
    headers = [cell.value for cell in ws[5]]
    data = [list(row) for row in ws.iter_rows(min_row=6, values_only=True)]
    df = pd.DataFrame(data, columns=headers)
    df.columns = [norm(c) if c else "" for c in df.columns]
    return df


def resolve_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def get_cols(df):
    return {
        "carrier":       resolve_col(df, ["carrier"]),
        "xdock":         resolve_col(df, ["xdock"]),
        "folio":         resolve_col(df, ["folio almacen origen"]),
        "estatus":       resolve_col(df, ["estatus"]),
        "est_salida":    resolve_col(df, ["estatus salida", "estatus de salida", "status salida"]),
        "fecha_salida":  resolve_col(df, ["fecha de salida"]),
        "id_sitio":      resolve_col(df, ["id sitio"]),
        "nombre_sitio":  resolve_col(df, ["nombre de sitio"]),
        "no_pallet":     resolve_col(df, ["no. de pallet", "no de pallet"]),
        "tipo_pallet":   resolve_col(df, ["tipo de pallet"]),
        "tipo_material": resolve_col(df, ["tipo de material"]),
        "desc_material": resolve_col(df, ["descripcion material"]),
    }


# ─────────────────────────────────────────────────────────────────────────────
#  PIPELINE STEPS
# ─────────────────────────────────────────────────────────────────────────────
def filter_salidas(df, cols):
    mask = (
        df[cols["estatus"]].apply(lambda x: norm(x) == "salida")
        | df[cols["est_salida"]].apply(lambda x: norm(x) == "salida")
        | df[cols["fecha_salida"]].notna()
    )
    return df[~mask].copy(), df[mask].copy()


def normalize_xdock_name(val):
    return XDOCK_ALIASES.get(norm(val), val)


def filter_ens(df, cols):
    df[cols["xdock"]] = df[cols["xdock"]].apply(
        lambda x: normalize_xdock_name(x) if not is_empty_xdock(x) else x
    )
    mask_ens   = df[cols["xdock"]].apply(lambda x: contains_ens(x))
    mask_empty = df[cols["xdock"]].apply(lambda x: is_empty_xdock(x))
    keep = mask_ens | mask_empty
    return df[keep].copy(), df[~keep].copy()


def infer_xdock_from_folio(folio):
    if not folio:
        return None
    f = str(folio).strip().upper()
    for prefix, xd in FOLIO_XDOCK:
        if f.startswith(prefix.upper()):
            return xd
    return None


def infer_xdock_from_id(id_sitio):
    if not id_sitio:
        return None
    s = str(id_sitio).strip().upper()
    if s == "HMTY0150":
        return "__DELETE__"
    if "JUMPER" in s.lower():
        return "__DELETE__"
    for prefix, xd in ID_SITIO_PREFIXES.items():
        if s.startswith(prefix.upper()):
            return xd
    return None


def infer_xdock_from_keywords(*fields):
    combined = " ".join(norm(str(f)) for f in fields if f)
    if "devolucion logistica inversa la paz" in combined:
        return "Gaso- La Paz-E-NS"
    for kw, xd in KEYWORD_XDOCK.items():
        if kw in combined:
            return xd
    return None


def fill_xdock(df, cols, saved_decisions: dict):
    """
    Fill empty XDOCK values.
    saved_decisions: {id_sitio_key -> {"action": "assign"|"delete", "xdock": "..."}}
    Returns (df_active, df_pending, n_auto_assigned, n_auto_deleted, n_decision_applied)
    """
    delete_idx  = []
    pending_idx = []
    n_auto  = 0
    n_saved = 0

    for idx, row in df.iterrows():
        xd = row[cols["xdock"]]
        if not is_empty_xdock(xd):
            continue

        folio     = row[cols["folio"]]
        id_sitio  = row[cols["id_sitio"]]
        nom_sitio = row[cols["nombre_sitio"]]
        desc      = row.get(cols["desc_material"], "")
        key       = str(id_sitio).strip().upper() if id_sitio else "__NOKEY__"

        # ── Check saved decision first ────────────────────────────────────
        if key in saved_decisions:
            dec = saved_decisions[key]
            if dec["action"] == "delete":
                delete_idx.append(idx)
                n_saved += 1
                continue
            elif dec["action"] == "assign" and dec.get("xdock"):
                df.at[idx, cols["xdock"]] = dec["xdock"]
                n_saved += 1
                continue

        # ── Auto-inference ────────────────────────────────────────────────
        result = infer_xdock_from_folio(folio)
        if result is None:
            result = infer_xdock_from_id(id_sitio)
        if result is None:
            result = infer_xdock_from_keywords(id_sitio, nom_sitio, folio, desc)

        if result == "__DELETE__":
            delete_idx.append(idx)
            n_auto += 1
        elif result:
            df.at[idx, cols["xdock"]] = result
            n_auto += 1
        else:
            pending_idx.append(idx)

    df = df.drop(index=delete_idx)
    df_pending = df.loc[df.index.isin(pending_idx)].copy()
    df_active  = df[~df.index.isin(pending_idx)].copy()
    return df_active, df_pending, n_auto, n_saved


def final_ens_filter(df, cols):
    mask = df[cols["xdock"]].apply(lambda x: contains_ens(str(x)))
    return df[mask].copy(), df[~mask].copy()


def fix_pallets(df, cols):
    def calc(row):
        if str(row[cols["id_sitio"]]).strip().upper() == "JALZAP0981":
            return 14
        val = row[cols["no_pallet"]]
        if pd.isna(val) or val == "" or val is None:
            return 1
        try:
            return 1  # any non-zero value → 1
        except Exception:
            return 1
    df[cols["no_pallet"]] = df.apply(calc, axis=1)
    return df


def normalize_tipo_pallet(df, cols):
    def fix(val):
        if pd.isna(val) or val == "" or val is None:
            return "ESTANDAR"
        n = norm(val)
        for k, v in PALLET_NORM.items():
            if k in n:
                return v
        return "ESTANDAR"
    df[cols["tipo_pallet"]] = df[cols["tipo_pallet"]].apply(fix)
    return df


def calc_m2(df, cols):
    def m2(row):
        tp  = str(row[cols["tipo_pallet"]]).upper().strip()
        np_ = float(row[cols["no_pallet"]]) if row[cols["no_pallet"]] else 0
        return round(np_ * M2_TIPO.get(tp, 1.44) * 1.20, 4)
    df["M2"] = df.apply(m2, axis=1)
    return df


def assign_region(df, cols):
    df["REGION"] = df[cols["xdock"]].map(REGION_MAP).fillna("SIN REGIÓN")
    df["CIUDAD"] = df[cols["xdock"]].map(CIUDAD_MAP).fillna("")
    return df


# ─────────────────────────────────────────────────────────────────────────────
#  MASTER PIPELINE
# ─────────────────────────────────────────────────────────────────────────────
def run_pipeline(uploaded_file, saved_decisions: dict):
    logs = []
    df_raw = load_file(uploaded_file)
    logs.append(f"✅ Archivo cargado: {len(df_raw):,} registros totales")

    cols = get_cols(df_raw)

    df, rem_sal = filter_salidas(df_raw, cols)
    logs.append(f"🗑️  Salidas eliminadas: {len(rem_sal):,} | Restantes: {len(df):,}")

    df, rem_ens = filter_ens(df, cols)
    logs.append(f"🗑️  Sin E-NS eliminados: {len(rem_ens):,} | Restantes: {len(df):,}")

    df, df_pending, n_auto, n_saved = fill_xdock(df, cols, saved_decisions)
    logs.append(f"🔍 XDOCK resueltos automáticamente: {n_auto} | Por decisión guardada: {n_saved} | Pendientes nuevos: {len(df_pending)}")

    df, rem_fin = final_ens_filter(df, cols)
    logs.append(f"🗑️  Filtro final E-NS: {len(rem_fin)} eliminados | Activos: {len(df):,}")

    df = fix_pallets(df, cols)
    df = normalize_tipo_pallet(df, cols)
    df = calc_m2(df, cols)
    df = assign_region(df, cols)
    logs.append(f"📦 Pallets: {int(df[cols['no_pallet']].sum()):,} | M²: {df['M2'].sum():,.2f}")

    return df, df_pending, logs, cols


# ─────────────────────────────────────────────────────────────────────────────
#  REPORT BUILDERS
# ─────────────────────────────────────────────────────────────────────────────
def build_client_report(df, cols):
    grp = df.groupby([cols["xdock"], cols["carrier"], cols["tipo_material"]]).agg(
        NO_PALLETS=(cols["no_pallet"], "sum"),
        M2=("M2", "sum"),
    ).reset_index()
    grp["CIUDAD"]    = grp[cols["xdock"]].map(CIUDAD_MAP).fillna(grp[cols["xdock"]])
    grp["REGION"]    = grp[cols["xdock"]].map(REGION_MAP).fillna("SIN REGIÓN")
    grp["CAPACIDAD"] = grp[cols["xdock"]].map(CAPACIDADES).fillna(0)
    grp["PCT_OCP"]   = grp.apply(
        lambda r: round(r["M2"] / r["CAPACIDAD"], 4) if r["CAPACIDAD"] > 0 else 0, axis=1
    )
    return grp


def build_pivot_m2(df, cols):
    pivot = df.pivot_table(
        index=cols["xdock"], columns=cols["tipo_material"],
        values="M2", aggfunc="sum", fill_value=0
    )
    pivot["TOTAL M2"]    = pivot.sum(axis=1)
    pivot["CAPACIDAD"]   = pivot.index.map(CAPACIDADES)
    pivot["% OCUPACIÓN"] = (pivot["TOTAL M2"] / pivot["CAPACIDAD"]).round(4)
    return pivot


def build_pivot_pallets(df, cols):
    pivot = df.pivot_table(
        index=cols["xdock"], columns=cols["tipo_material"],
        values=cols["no_pallet"], aggfunc="sum", fill_value=0
    )
    pivot["TOTAL PALLETS"] = pivot.sum(axis=1)
    return pivot


# ─────────────────────────────────────────────────────────────────────────────
#  EXCEL EXPORT  (4 sheets: IN-OUT LIMPIO · REPORTE CLIENTE · PIVOT M2 · RESUMEN)
# ─────────────────────────────────────────────────────────────────────────────
HEX_BLUE  = "1A3A6B"
HEX_LIGHT = "2E6DB4"
HEX_WHITE = "FFFFFF"
HEX_LGRAY = "EBF0F7"
HEX_GREEN = "1E8449"
HEX_AMBER = "E67E22"
HEX_RED   = "C0392B"
HEX_DRED  = "7B241C"

_thin = Side(style="thin", color="BBBBBB")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

HDR_FONT   = Font(name="Calibri", bold=True, color=HEX_WHITE, size=10)
HDR_FILL   = PatternFill("solid", fgColor=HEX_BLUE)
DATA_FONT  = Font(name="Calibri", size=9)
CTR        = Alignment(horizontal="center", vertical="center", wrap_text=True)
LFT        = Alignment(horizontal="left",   vertical="center")


def _pct_fill(pct):
    if pct > 1.00: return PatternFill("solid", fgColor=HEX_DRED)
    if pct > 0.90: return PatternFill("solid", fgColor=HEX_RED)
    if pct > 0.70: return PatternFill("solid", fgColor=HEX_AMBER)
    return PatternFill("solid", fgColor=HEX_GREEN)


def _alt_fill(i):
    return PatternFill("solid", fgColor=HEX_LGRAY if i % 2 else HEX_WHITE)


def _title_block(ws, title, subtitle, fecha, n_cols=18):
    ws.row_dimensions[1].height = 34
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 16
    end = get_column_letter(n_cols)
    for row, val, size, fg, bg in [
        (1, "GASO COMUNICACIONES",  15, HEX_WHITE, HEX_BLUE),
        (2, title,                  11, HEX_WHITE, HEX_LIGHT),
        (3, f"{subtitle}   |   Fecha generación: {fecha}", 9, "555555", "F4F6F9"),
    ]:
        ws.merge_cells(f"A{row}:{end}{row}")
        c = ws[f"A{row}"]
        c.value     = val
        c.font      = Font(name="Calibri", bold=(row < 3), size=size, color=fg)
        c.fill      = PatternFill("solid", fgColor=bg)
        c.alignment = CTR


def _write_table(ws, df, start_row, freeze=True):
    cols = list(df.columns)
    for ci, h in enumerate(cols, 1):
        c = ws.cell(row=start_row, column=ci, value=str(h))
        c.font = HDR_FONT; c.fill = HDR_FILL
        c.border = BORDER; c.alignment = CTR
    for ri, (_, row) in enumerate(df.iterrows()):
        er = start_row + 1 + ri
        for ci, val in enumerate(row, 1):
            c = ws.cell(row=er, column=ci, value=val)
            c.font = DATA_FONT; c.border = BORDER
            c.fill = _alt_fill(ri); c.alignment = LFT
    # auto-width
    for ci, col in enumerate(cols, 1):
        try:
            mx = max(len(str(col)),
                     df.iloc[:, ci-1].astype(str).str.len().max() if len(df) else 10)
        except Exception:
            mx = 14
        ws.column_dimensions[get_column_letter(ci)].width = min(mx + 3, 42)
    if freeze:
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)


def build_excel_output(df_clean, pivot_m2, pivot_pal, cols):
    wb   = openpyxl.Workbook()
    fecha = datetime.date.today().strftime("%d/%m/%Y")

    xdock_col  = cols["xdock"]
    carrier_col = cols["carrier"]
    mat_col    = cols["tipo_material"]
    pallet_col = cols["no_pallet"]

    # ── 1. IN-OUT LIMPIO ────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "IN-OUT LIMPIO"
    _title_block(ws1, "BASE DE DATOS – IN-OUT LIMPIO",
                 "Inventario activo depurado y enriquecido", fecha, n_cols=min(30, len(df_clean.columns)))
    _write_table(ws1, df_clean.reset_index(drop=True), start_row=5)

    # ── 2. REPORTE CLIENTE ──────────────────────────────────────────────────
    ws2 = wb.create_sheet("REPORTE CLIENTE")
    ws2.sheet_view.showGridLines = False

    carriers = sorted(df_clean[carrier_col].dropna().unique())
    tipos    = sorted(df_clean[mat_col].dropna().unique())
    xdocks   = sorted(CAPACIDADES.keys())

    # Build header columns
    exec_cols  = ["CIUDAD", "REGIÓN", "CAPACIDAD M²"]
    for car in carriers:
        for tp in tipos:
            exec_cols.append(f"{car}\n{tp[:18]}\nPallets")
        exec_cols.append(f"{car}\nTotal Pallets")
        exec_cols.append(f"{car}\nM² Ocupados")
    exec_cols += ["TOTAL\nPALLETS", "TOTAL M²\nOCUPADOS", "% OCUPACIÓN", "DISPONIBLE M²", "STATUS"]

    HDR_ROW = 5
    _title_block(ws2, "REPORTE SEMANAL DE OCUPACIÓN – CLIENTE",
                 "Resumen ejecutivo por XDOCK · Carrier · Tipo de Material", fecha, n_cols=len(exec_cols))

    ws2.row_dimensions[HDR_ROW].height = 40
    for ci, h in enumerate(exec_cols, 1):
        c = ws2.cell(row=HDR_ROW, column=ci, value=h)
        c.font = HDR_FONT; c.fill = HDR_FILL
        c.border = BORDER; c.alignment = CTR
        ws2.column_dimensions[get_column_letter(ci)].width = 13
    ws2.column_dimensions["A"].width = 14
    ws2.column_dimensions["B"].width = 14
    ws2.column_dimensions["C"].width = 12

    total_pal_all = 0
    total_m2_all  = 0.0

    for ri, xd in enumerate(xdocks):
        er  = HDR_ROW + 1 + ri
        ciudad = CIUDAD_MAP.get(xd, xd)
        region = REGION_MAP.get(xd, "")
        cap    = CAPACIDADES.get(xd, 0)

        row_vals = [ciudad, region, cap]
        tot_pal = 0; tot_m2 = 0.0

        for car in carriers:
            sub = df_clean[(df_clean[xdock_col] == xd) & (df_clean[carrier_col] == car)]
            for tp in tipos:
                sub2 = sub[sub[mat_col] == tp]
                row_vals.append(int(sub2[pallet_col].sum()))
            c_pal = int(sub[pallet_col].sum())
            c_m2  = round(sub["M2"].sum(), 2)
            row_vals += [c_pal, c_m2]
            tot_pal += c_pal; tot_m2 += c_m2

        tot_m2 = round(tot_m2, 2)
        pct    = round(tot_m2 / cap, 4) if cap > 0 else 0
        disp   = round(cap - tot_m2, 2)
        status = ("SATURADO" if pct > 1.0 else "CRÍTICO" if pct > 0.90
                  else "ALERTA" if pct > 0.70 else "NORMAL")
        row_vals += [tot_pal, tot_m2, pct, disp, status]

        total_pal_all += tot_pal
        total_m2_all  += tot_m2

        n_fixed_tail = 5   # TOTAL PALLETS · TOTAL M2 · %OCP · DISPONIBLE · STATUS
        pct_col_idx  = len(exec_cols) - 2   # 0-based distance from end

        for ci, val in enumerate(row_vals, 1):
            c = ws2.cell(row=er, column=ci, value=val)
            c.border = BORDER; c.font = DATA_FONT
            c.fill   = _alt_fill(ri); c.alignment = CTR

        # Color % ocupación
        n_cols_total = len(exec_cols)
        pct_ci = n_cols_total - 2   # 1-based
        c_pct  = ws2.cell(row=er, column=pct_ci)
        c_pct.number_format = "0.00%"
        c_pct.fill = _pct_fill(pct)
        c_pct.font = Font(name="Calibri", bold=True, color=HEX_WHITE, size=9)
        c_pct.alignment = CTR

    # Totals row
    tr = HDR_ROW + 1 + len(xdocks)
    total_cap = sum(CAPACIDADES.values())
    pct_global = round(total_m2_all / total_cap, 4)

    ws2.cell(row=tr, column=1, value="TOTAL GLOBAL").font  = Font(name="Calibri", bold=True, color=HEX_WHITE, size=10)
    ws2.cell(row=tr, column=1).fill  = PatternFill("solid", fgColor=HEX_BLUE)
    ws2.cell(row=tr, column=1).alignment = CTR
    ws2.cell(row=tr, column=3, value=total_cap).font  = Font(name="Calibri", bold=True, color=HEX_WHITE, size=9)
    ws2.cell(row=tr, column=3).fill = PatternFill("solid", fgColor=HEX_BLUE)
    ws2.cell(row=tr, column=3).alignment = CTR

    nc = len(exec_cols)
    for ci, val in [(nc-3, total_pal_all), (nc-2, round(total_m2_all, 2))]:
        c = ws2.cell(row=tr, column=ci, value=val)
        c.font = Font(name="Calibri", bold=True, color=HEX_WHITE, size=9)
        c.fill = PatternFill("solid", fgColor=HEX_BLUE); c.alignment = CTR

    c_pg = ws2.cell(row=tr, column=nc-1, value=pct_global)
    c_pg.number_format = "0.00%"
    c_pg.fill = _pct_fill(pct_global)
    c_pg.font = Font(name="Calibri", bold=True, color=HEX_WHITE, size=9)
    c_pg.alignment = CTR

    ws2.freeze_panes = ws2.cell(row=HDR_ROW + 1, column=1)

    # ── 3. PIVOT M2 ──────────────────────────────────────────────────────────
    ws3 = wb.create_sheet("PIVOT M2")
    _title_block(ws3, "PIVOT – M² POR CROSSDOCK Y TIPO DE MATERIAL",
                 "Análisis de metros cuadrados · Factor pasillos 1.20x incluido", fecha)
    pm2_out = pivot_m2.copy().reset_index()
    _write_table(ws3, pm2_out, start_row=5)
    # color % ocupacion column
    if "% OCUPACIÓN" in pm2_out.columns:
        pci = list(pm2_out.columns).index("% OCUPACIÓN") + 1
        for ri in range(len(pm2_out)):
            c = ws3.cell(row=6 + ri, column=pci)
            c.number_format = "0.00%"
            try:
                c.fill = _pct_fill(float(c.value))
                c.font = Font(name="Calibri", bold=True, color=HEX_WHITE, size=9)
                c.alignment = CTR
            except Exception:
                pass

    # ── 4. RESUMEN EJECUTIVO ─────────────────────────────────────────────────
    ws4 = wb.create_sheet("RESUMEN EJECUTIVO")
    ws4.sheet_view.showGridLines = False
    _title_block(ws4, "RESUMEN EJECUTIVO DE OCUPACIÓN",
                 "Indicadores clave de desempeño del inventario en tiempo real", fecha, n_cols=8)

    total_pal   = int(df_clean[pallet_col].sum())
    total_m2    = round(df_clean["M2"].sum(), 2)
    disp_global = round(total_cap - total_m2, 2)

    # KPI block
    kpis = [
        ("Total Pallets en Inventario",   f"{total_pal:,}",         "unidades"),
        ("Total M² Ocupados",             f"{total_m2:,.2f}",        "m²"),
        ("Capacidad Total del Sistema",   f"{total_cap:,}",          "m²"),
        ("% Ocupación Global",            f"{pct_global*100:.1f}%",  ""),
        ("M² Disponibles",                f"{disp_global:,.2f}",     "m²"),
        ("Crossdocks activos",            f"{df_clean[xdock_col].nunique()}", "xdocks"),
    ]
    for ci, h in enumerate(["INDICADOR", "VALOR", "UNIDAD"], 1):
        c = ws4.cell(row=5, column=ci, value=h)
        c.font = HDR_FONT; c.fill = HDR_FILL; c.border = BORDER; c.alignment = CTR
    ws4.column_dimensions["A"].width = 32
    ws4.column_dimensions["B"].width = 18
    ws4.column_dimensions["C"].width = 12
    for ri, (kpi, val, unit) in enumerate(kpis, 6):
        f = _alt_fill(ri)
        for ci, v in enumerate([kpi, val, unit], 1):
            c = ws4.cell(row=ri, column=ci, value=v)
            c.fill = f; c.border = BORDER
            c.font = Font(name="Calibri", bold=(ci == 1), size=10)
            c.alignment = CTR if ci > 1 else LFT

    # Occupancy table
    ws4.cell(row=13, column=1, value="OCUPACIÓN POR CROSSDOCK").font = \
        Font(name="Calibri", bold=True, size=11, color=HEX_BLUE)
    ocp_hdr = ["CIUDAD", "REGIÓN", "CAP. M²", "PALLETS", "M² OCUPADOS", "% OCUPACIÓN", "DISPONIBLE M²", "STATUS"]
    for ci, h in enumerate(ocp_hdr, 1):
        c = ws4.cell(row=14, column=ci, value=h)
        c.font = HDR_FONT; c.fill = HDR_FILL; c.border = BORDER; c.alignment = CTR
        ws4.column_dimensions[get_column_letter(ci)].width = 16

    for ri, xd in enumerate(xdocks, 15):
        ciudad = CIUDAD_MAP.get(xd, xd)
        cap    = CAPACIDADES.get(xd, 0)
        m2_ocp = round(df_clean[df_clean[xdock_col] == xd]["M2"].sum(), 2)
        pal    = int(df_clean[df_clean[xdock_col] == xd][pallet_col].sum())
        pct    = round(m2_ocp / cap, 4) if cap > 0 else 0
        disp   = round(cap - m2_ocp, 2)
        region = REGION_MAP.get(xd, "")
        status = ("SATURADO" if pct > 1.0 else "CRÍTICO" if pct > 0.90
                  else "ALERTA" if pct > 0.70 else "NORMAL")
        row_v = [ciudad, region, cap, pal, m2_ocp, pct, disp, status]
        f = _alt_fill(ri)
        for ci, val in enumerate(row_v, 1):
            c = ws4.cell(row=ri, column=ci, value=val)
            c.fill = f; c.border = BORDER; c.font = DATA_FONT; c.alignment = CTR
        c_pct = ws4.cell(row=ri, column=6)
        c_pct.number_format = "0.00%"
        c_pct.fill = _pct_fill(pct)
        c_pct.font = Font(name="Calibri", bold=True, color=HEX_WHITE, size=9)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────────────────────
#  PLOTLY CHARTS
# ─────────────────────────────────────────────────────────────────────────────
def _color_pct(pct):
    if pct > 1.0:  return "#7B241C"
    if pct > 0.90: return "#C0392B"
    if pct > 0.70: return "#E67E22"
    return "#1E8449"


def make_charts(df_clean, cols):
    xdock_col  = cols["xdock"]
    carrier_col = cols["carrier"]
    mat_col    = cols["tipo_material"]
    pallet_col = cols["no_pallet"]

    # Occupancy bar
    ocp = []
    for xd in CAPACIDADES:
        cap    = CAPACIDADES[xd]
        m2_ocp = df_clean[df_clean[xdock_col] == xd]["M2"].sum()
        pct    = m2_ocp / cap if cap > 0 else 0
        ocp.append({"Ciudad": CIUDAD_MAP.get(xd, xd), "pct": round(pct*100,1),
                    "M2_ocp": round(m2_ocp,1), "cap": cap, "disp": round(cap-m2_ocp,1)})
    df_ocp = pd.DataFrame(ocp).sort_values("pct", ascending=True)

    fig1 = go.Figure(go.Bar(
        x=df_ocp["pct"], y=df_ocp["Ciudad"], orientation="h",
        marker_color=[_color_pct(p/100) for p in df_ocp["pct"]],
        text=[f"{p}%" for p in df_ocp["pct"]], textposition="outside",
        hovertemplate="<b>%{y}</b><br>Ocupación: %{x}%<extra></extra>",
    ))
    for v, label, color in [(70, "70%", "#E67E22"), (90, "90%", "#C0392B"), (100, "100%", "#7B241C")]:
        fig1.add_vline(x=v, line_dash="dot", line_color=color,
                       annotation_text=label, annotation_font_color=color)
    fig1.update_layout(
        title="% Ocupación por Crossdock",
        xaxis=dict(range=[0, max(df_ocp["pct"].max()*1.15, 115)], title="% Ocupación"),
        yaxis_title="", plot_bgcolor="white", paper_bgcolor="white",
        height=380, margin=dict(l=10, r=60, t=40, b=20),
        font=dict(family="Calibri", color=GASO_BLUE),
    )

    # Stacked M2
    fig2 = go.Figure([
        go.Bar(x=df_ocp["Ciudad"], y=df_ocp["M2_ocp"], name="M² Ocupados",
               marker_color=GASO_LIGHT,
               hovertemplate="<b>%{x}</b><br>Ocupados: %{y} m²<extra></extra>"),
        go.Bar(x=df_ocp["Ciudad"], y=df_ocp["disp"], name="Disponible",
               marker_color="#D5E8F5",
               hovertemplate="<b>%{x}</b><br>Disponible: %{y} m²<extra></extra>"),
    ])
    fig2.update_layout(
        barmode="stack", title="M² Ocupados vs Disponibles",
        xaxis_title="", yaxis_title="m²", plot_bgcolor="white", paper_bgcolor="white",
        height=360, margin=dict(l=20, r=20, t=40, b=60),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        font=dict(family="Calibri", color=GASO_BLUE),
    )

    # Material bar
    mat = df_clean.groupby(mat_col)[pallet_col].sum().reset_index()
    mat.columns = ["Tipo", "Pallets"]
    mat = mat.sort_values("Pallets", ascending=False)
    fig3 = px.bar(mat, x="Tipo", y="Pallets", text="Pallets",
                  color="Tipo", color_discrete_sequence=px.colors.qualitative.Bold,
                  title="Pallets por Tipo de Material")
    fig3.update_traces(textposition="outside")
    fig3.update_layout(showlegend=False, plot_bgcolor="white", paper_bgcolor="white",
                       height=340, margin=dict(l=20, r=20, t=40, b=60),
                       font=dict(family="Calibri", color=GASO_BLUE))

    # Carrier pie
    car = df_clean.groupby(carrier_col)[pallet_col].sum().reset_index()
    car.columns = ["Carrier", "Pallets"]
    fig4 = px.pie(car, names="Carrier", values="Pallets", hole=0.45,
                  title="Pallets por Carrier",
                  color_discrete_sequence=[GASO_BLUE, GASO_ACCENT, "#E67E22", "#8E44AD"])
    fig4.update_traces(textinfo="percent+label", pull=[0.03]*len(car))
    fig4.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                       height=340, margin=dict(l=20, r=20, t=40, b=20),
                       font=dict(family="Calibri", color=GASO_BLUE))

    # Heatmap
    heat = df_clean.groupby([xdock_col, mat_col])["M2"].sum().reset_index()
    hp   = heat.pivot(index=xdock_col, columns=mat_col, values="M2").fillna(0)
    hp.index = [CIUDAD_MAP.get(x, x) for x in hp.index]
    fig5 = px.imshow(hp.round(1), text_auto=".0f", aspect="auto",
                     color_continuous_scale=[[0,"#EBF5FB"],[0.5,GASO_ACCENT],[1,GASO_BLUE]],
                     title="Heatmap M²: Crossdock × Tipo de Material")
    fig5.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                       height=380, margin=dict(l=10, r=10, t=50, b=10),
                       font=dict(family="Calibri", color=GASO_BLUE))

    # Region bar
    reg = df_clean.groupby("REGION")["M2"].sum().reset_index()
    fig6 = px.bar(reg, x="REGION", y="M2", text="M2", color="REGION",
                  color_discrete_map={"REGIÓN LUIS": GASO_BLUE, "REGIÓN JORGE": GASO_ACCENT},
                  title="M² por Región")
    fig6.update_traces(texttemplate="%{text:.0f} m²", textposition="outside")
    fig6.update_layout(showlegend=False, plot_bgcolor="white", paper_bgcolor="white",
                       height=320, margin=dict(l=20, r=20, t=40, b=20),
                       font=dict(family="Calibri", color=GASO_BLUE))

    # Carrier × Ciudad
    cx = df_clean.groupby([carrier_col, xdock_col])[pallet_col].sum().reset_index()
    cx["Ciudad"] = cx[xdock_col].map(CIUDAD_MAP)
    fig7 = px.bar(cx, x="Ciudad", y=pallet_col, color=carrier_col, barmode="group",
                  color_discrete_sequence=[GASO_BLUE, GASO_ACCENT, "#E67E22"],
                  title="Pallets por Ciudad y Carrier")
    fig7.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                       height=320, margin=dict(l=20, r=20, t=40, b=60),
                       font=dict(family="Calibri", color=GASO_BLUE),
                       xaxis_title="", yaxis_title="Pallets",
                       legend=dict(title="Carrier"))

    return fig1, fig2, fig3, fig4, fig5, fig6, fig7


# ─────────────────────────────────────────────────────────────────────────────
#  STREAMLIT APP
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="GASO – IN-OUT Processor",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(f"""
<style>
html, body, [class*="css"] {{ font-family: 'Calibri', 'Segoe UI', sans-serif; }}
.main-header {{
    background: linear-gradient(135deg, {GASO_BLUE} 0%, {GASO_LIGHT} 100%);
    padding: 1.4rem 2rem; border-radius: 12px; margin-bottom: 1.2rem;
    display: flex; align-items: center; gap: 1.2rem;
}}
.main-header h1 {{ color: white; margin: 0; font-size: 1.7rem; font-weight: 700; }}
.main-header p  {{ color: rgba(255,255,255,.82); margin: 0; font-size: 0.88rem; }}
.kpi-card {{
    background: white; border: 1px solid #DCE8F5;
    border-left: 4px solid {GASO_BLUE};
    border-radius: 10px; padding: .9rem 1rem;
    box-shadow: 0 2px 8px rgba(26,58,107,.08); text-align: center;
}}
.kpi-label {{ color:#666; font-size:.72rem; font-weight:700;
              text-transform:uppercase; letter-spacing:.05em; margin-bottom:.3rem; }}
.kpi-value {{ color:{GASO_BLUE}; font-size:1.75rem; font-weight:700; line-height:1; }}
.kpi-unit  {{ color:#999; font-size:.72rem; margin-top:.2rem; }}
.kpi-red   {{ border-left-color:#C0392B; }}
.kpi-amber {{ border-left-color:#E67E22; }}
.kpi-green {{ border-left-color:#1E8449; }}
.sec-title {{
    color:{GASO_BLUE}; font-size:1.05rem; font-weight:700;
    border-bottom: 2px solid {GASO_ACCENT};
    padding-bottom:.25rem; margin: 1.1rem 0 .7rem 0;
}}
.review-card {{
    background:#FFF8F0; border:1px solid #F0C080;
    border-left:4px solid #E67E22; border-radius:10px;
    padding:1rem 1.2rem; margin-bottom:.8rem;
}}
.review-card h4 {{ color:#7D3C00; margin:0 0 .3rem 0; font-size:.95rem; }}
.review-card p  {{ color:#555; font-size:.82rem; margin:0; }}
.saved-badge {{
    background:#D5F5E3; color:#1E8449; border-radius:20px;
    padding:.15rem .6rem; font-size:.75rem; font-weight:700;
}}
.del-badge {{
    background:#FADBD8; color:#C0392B; border-radius:20px;
    padding:.15rem .6rem; font-size:.75rem; font-weight:700;
}}
.log-box {{
    background:#F0F4F8; border-radius:8px; padding:.7rem 1rem;
    font-size:.8rem; font-family:monospace;
    max-height:200px; overflow-y:auto;
}}
.stDownloadButton > button {{
    background:{GASO_BLUE}; color:white; border:none;
    border-radius:8px; padding:.55rem 1.2rem;
    font-weight:600; font-size:.88rem; width:100%;
    transition:background .2s;
}}
.stDownloadButton > button:hover {{ background:{GASO_LIGHT}; }}
</style>
""", unsafe_allow_html=True)

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center;padding:.8rem 0">
      <div style="background:{GASO_BLUE};color:white;border-radius:50%;
                  width:56px;height:56px;line-height:56px;font-size:1.6rem;
                  margin:0 auto .4rem auto;font-weight:700;">G</div>
      <div style="color:{GASO_BLUE};font-weight:700;font-size:.95rem;">GASO COMUNICACIONES</div>
      <div style="color:#888;font-size:.72rem;">IN-OUT Report Processor</div>
    </div><hr style="border-color:#DCE8F5">
    """, unsafe_allow_html=True)

    st.markdown("### 📁 Cargar Archivo")
    uploaded = st.file_uploader(
        "Selecciona el archivo IN-OUT (.xlsx)",
        type=["xlsx"],
        help="Hoja 'IN-OUT' con encabezados en fila 5."
    )

    st.markdown("---")
    st.markdown("### ⚙️ Opciones de Vista")
    show_clean = st.checkbox("Mostrar base limpia", value=False)
    show_piv   = st.checkbox("Mostrar pivots",      value=False)

    # Decision memory manager
    st.markdown("---")
    st.markdown("### 🧠 Decisiones Guardadas")
    saved = load_decisions()
    if saved:
        st.caption(f"{len(saved)} ID_SITIO(s) con decisión guardada")
        for key, dec in saved.items():
            col_a, col_b = st.columns([3, 1])
            with col_a:
                label = dec.get("xdock", "ELIMINAR") if dec["action"] == "assign" else "🗑 ELIMINAR"
                st.caption(f"**{key}** → {label}")
            with col_b:
                if st.button("✕", key=f"del_{key}", help="Borrar decisión"):
                    delete_decision(key)
                    st.rerun()
    else:
        st.caption("Sin decisiones guardadas aún.")

    st.markdown("---")
    st.markdown(f"""
    <div style="font-size:.7rem;color:#AAA;text-align:center">
    v2.0.0 · Gaso Comunicaciones<br>Procesamiento automático de inventario
    </div>""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="main-header">
  <div style="font-size:2.2rem">📦</div>
  <div>
    <h1>IN-OUT Report Processor</h1>
    <p>Limpieza · Enriquecimiento · Reporte Ejecutivo para Cliente &nbsp;|&nbsp; Gaso Comunicaciones</p>
  </div>
</div>
""", unsafe_allow_html=True)

if not uploaded:
    st.markdown("""
    <div style="text-align:center;padding:3rem;background:#F8FAFC;
                border-radius:12px;border:2px dashed #C8D8EC">
      <div style="font-size:3rem">📂</div>
      <h3 style="color:#1A3A6B">Carga tu archivo IN-OUT para comenzar</h3>
      <p style="color:#888">Formato esperado: <b>Excel (.xlsx)</b> · Hoja <b>IN-OUT</b> · Encabezados en fila 5</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── Process button ────────────────────────────────────────────────────────────
_, col_btn, _ = st.columns([2, 1, 2])
with col_btn:
    process_btn = st.button("🚀 Procesar Archivo", type="primary", use_container_width=True)

if "processed" not in st.session_state:
    st.session_state.processed = False

if process_btn:
    uploaded.seek(0)
    with st.spinner("⚙️ Procesando base de datos..."):
        try:
            saved_dec = load_decisions()
            df_clean, df_pending, logs, cols = run_pipeline(uploaded, saved_dec)

            pivot_m2  = build_pivot_m2(df_clean, cols)
            pivot_pal = build_pivot_pallets(df_clean, cols)

            st.session_state.df_clean   = df_clean
            st.session_state.df_pending = df_pending
            st.session_state.pivot_m2   = pivot_m2
            st.session_state.pivot_pal  = pivot_pal
            st.session_state.logs       = logs
            st.session_state.cols       = cols
            st.session_state.processed  = True

            # Pre-build Excel
            uploaded.seek(0)
            excel_buf = build_excel_output(df_clean, pivot_m2, pivot_pal, cols)
            st.session_state.excel_buf = excel_buf

        except Exception as e:
            st.error(f"❌ Error durante el procesamiento: {e}")
            import traceback; st.code(traceback.format_exc())

# ── Results ───────────────────────────────────────────────────────────────────
if st.session_state.processed:
    df_clean   = st.session_state.df_clean
    df_pending = st.session_state.df_pending
    pivot_m2   = st.session_state.pivot_m2
    pivot_pal  = st.session_state.pivot_pal
    logs       = st.session_state.logs
    cols       = st.session_state.cols
    excel_buf  = st.session_state.excel_buf

    xdock_col  = cols["xdock"]
    pallet_col = cols["no_pallet"]

    # ── Processing log ────────────────────────────────────────────────────────
    with st.expander("📋 Log de Procesamiento", expanded=False):
        st.markdown('<div class="log-box">' + "<br>".join(logs) + "</div>",
                    unsafe_allow_html=True)

    # ── KPIs ──────────────────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">📊 Indicadores Clave de Desempeño</div>',
                unsafe_allow_html=True)

    total_pal  = int(df_clean[pallet_col].sum())
    total_m2   = round(df_clean["M2"].sum(), 2)
    total_cap  = sum(CAPACIDADES.values())
    pct_global = round(total_m2 / total_cap * 100, 1)
    disponible = round(total_cap - total_m2, 2)
    n_pend     = len(df_pending)

    kpi_cls = "kpi-red" if pct_global > 90 else "kpi-amber" if pct_global > 70 else "kpi-green"

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    kpis_ui = [
        (k1, "Total Pallets",       f"{total_pal:,}",      "unidades",  "kpi-card"),
        (k2, "M² Ocupados",         f"{total_m2:,.1f}",    "m²",        "kpi-card"),
        (k3, "Capacidad Total",     f"{total_cap:,}",      "m²",        "kpi-card"),
        (k4, "% Ocupación Global",  f"{pct_global}%",      "",          f"kpi-card {kpi_cls}"),
        (k5, "M² Disponibles",      f"{disponible:,.1f}",  "m²",        "kpi-card"),
        (k6, "Pendientes Revisión", f"{n_pend}",            "registros",
         "kpi-card kpi-red" if n_pend > 0 else "kpi-card kpi-green"),
    ]
    for col_w, label, val, unit, cls in kpis_ui:
        with col_w:
            st.markdown(f"""
            <div class="{cls}">
              <div class="kpi-label">{label}</div>
              <div class="kpi-value">{val}</div>
              <div class="kpi-unit">{unit}</div>
            </div>""", unsafe_allow_html=True)

    # ── Occupancy table ───────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">🏭 Ocupación por Crossdock</div>',
                unsafe_allow_html=True)
    ocp_rows = []
    for xd in CAPACIDADES:
        cap    = CAPACIDADES[xd]
        m2_ocp = round(df_clean[df_clean[xdock_col] == xd]["M2"].sum(), 2)
        pal    = int(df_clean[df_clean[xdock_col] == xd][pallet_col].sum())
        pct    = round(m2_ocp / cap * 100, 1) if cap > 0 else 0
        ocp_rows.append({
            "Ciudad":         CIUDAD_MAP.get(xd, xd),
            "Región":         REGION_MAP.get(xd, ""),
            "Capacidad m²":   cap,
            "Pallets":        pal,
            "M² Ocupados":    m2_ocp,
            "% Ocupación":    pct,
            "Disponible m²":  round(cap - m2_ocp, 2),
            "Status":         ("🔴 SATURADO" if pct > 100 else "🔴 CRÍTICO" if pct > 90
                               else "🟡 ALERTA" if pct > 70 else "🟢 NORMAL"),
        })
    df_ocp_t = pd.DataFrame(ocp_rows)

    def _color_ocp_col(col):
        styles = []
        for v in col:
            try:
                p = float(v)
            except (TypeError, ValueError):
                styles.append("")
                continue
            if p > 100:
                styles.append("background-color:#7B241C;color:white;font-weight:700")
            elif p > 90:
                styles.append("background-color:#C0392B;color:white;font-weight:700")
            elif p > 70:
                styles.append("background-color:#E67E22;color:white;font-weight:700")
            else:
                styles.append("background-color:#1E8449;color:white;font-weight:700")
        return styles

    st.dataframe(
        df_ocp_t.style
            .apply(_color_ocp_col, subset=["% Ocupación"])
            .format({"% Ocupación": "{:.1f}%", "M² Ocupados": "{:,.2f}",
                     "Disponible m²": "{:,.2f}", "Capacidad m²": "{:,}"}),
        use_container_width=True, height=390,
    )

    # ── Charts ────────────────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">📈 Dashboard Visual</div>', unsafe_allow_html=True)
    fig1, fig2, fig3, fig4, fig5, fig6, fig7 = make_charts(df_clean, cols)

    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(fig1, use_container_width=True)
    with c2: st.plotly_chart(fig2, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3: st.plotly_chart(fig3, use_container_width=True)
    with c4: st.plotly_chart(fig4, use_container_width=True)

    st.plotly_chart(fig5, use_container_width=True)

    c5, c6 = st.columns(2)
    with c5: st.plotly_chart(fig6, use_container_width=True)
    with c6: st.plotly_chart(fig7, use_container_width=True)

    # ── Base Limpia (optional) ────────────────────────────────────────────────
    if show_clean:
        st.markdown('<div class="sec-title">🗂️ Base IN-OUT Limpia</div>',
                    unsafe_allow_html=True)
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            f_car = st.multiselect("Carrier",        df_clean[cols["carrier"]].dropna().unique().tolist())
        with fc2:
            f_xd  = st.multiselect("XDOCK",          df_clean[xdock_col].dropna().unique().tolist())
        with fc3:
            f_mat = st.multiselect("Tipo Material",  df_clean[cols["tipo_material"]].dropna().unique().tolist())
        df_show = df_clean.copy()
        if f_car: df_show = df_show[df_show[cols["carrier"]].isin(f_car)]
        if f_xd:  df_show = df_show[df_show[xdock_col].isin(f_xd)]
        if f_mat: df_show = df_show[df_show[cols["tipo_material"]].isin(f_mat)]
        st.markdown(f"**{len(df_show):,}** registros mostrados")
        st.dataframe(df_show.reset_index(drop=True), use_container_width=True, height=340)

    if show_piv:
        st.markdown('<div class="sec-title">📐 Pivot M²</div>', unsafe_allow_html=True)
        st.dataframe(pivot_m2.style.format("{:.2f}"), use_container_width=True, height=320)
        st.markdown('<div class="sec-title">📦 Pivot Pallets</div>', unsafe_allow_html=True)
        st.dataframe(pivot_pal, use_container_width=True, height=320)

    # ─────────────────────────────────────────────────────────────────────────
    #  REVISIÓN MANUAL  (with persistent memory)
    # ─────────────────────────────────────────────────────────────────────────
    if n_pend > 0:
        st.markdown(f"""
        <div class="sec-title">
          ⚠️ Revisión Manual de Registros
          &nbsp;<span style="background:#FDEBD0;color:#A04000;border-radius:20px;
          padding:.15rem .6rem;font-size:.78rem;font-weight:700">{n_pend} registros · {df_pending[cols['id_sitio']].nunique()} ID_SITIO únicos</span>
        </div>
        """, unsafe_allow_html=True)

        st.info(
            "Los registros siguientes no pudieron asignarse automáticamente. "
            "Toma una decisión por **ID Sitio** — se guardará y se aplicará "
            "automáticamente la próxima vez que subas el archivo. "
            "Puedes modificar o borrar decisiones en cualquier momento desde la barra lateral."
        )

        # Group by ID_SITIO
        pend_groups = (
            df_pending.groupby(cols["id_sitio"])
            .agg(
                REGISTROS=(cols["id_sitio"], "count"),
                CARRIER=(cols["carrier"], lambda x: ", ".join(x.dropna().unique())),
                NOMBRE_SITIO=(cols["nombre_sitio"], lambda x: ", ".join(x.dropna().unique()[:2])),
                FOLIO_EJEMPLO=(cols["folio"], "first"),
            )
            .reset_index()
        )

        current_decisions = load_decisions()
        any_new = False

        for _, grp_row in pend_groups.iterrows():
            id_s  = str(grp_row[cols["id_sitio"]]).strip()
            key   = id_s.upper()
            regs  = grp_row["REGISTROS"]
            carr  = grp_row["CARRIER"]
            nom   = grp_row["NOMBRE_SITIO"]
            folio = grp_row["FOLIO_EJEMPLO"]

            already = current_decisions.get(key, {})
            badge   = ""
            if already.get("action") == "assign":
                badge = f'<span class="saved-badge">✓ Guardado → {already["xdock"]}</span>'
            elif already.get("action") == "delete":
                badge = '<span class="del-badge">✓ Guardado → ELIMINAR</span>'

            st.markdown(f"""
            <div class="review-card">
              <h4>🔍 ID Sitio: <code>{id_s}</code> &nbsp; {badge}</h4>
              <p>Nombre: <b>{nom}</b> &nbsp;|&nbsp; Carrier: <b>{carr}</b>
                 &nbsp;|&nbsp; Registros afectados: <b>{regs}</b>
                 &nbsp;|&nbsp; Folio ejemplo: <code>{folio}</code></p>
            </div>
            """, unsafe_allow_html=True)

            col_dec, col_xd, col_save = st.columns([2, 3, 1])
            with col_dec:
                action = st.radio(
                    "Acción",
                    ["Asignar XDOCK", "Eliminar registros"],
                    key=f"action_{key}",
                    horizontal=True,
                    index=0 if already.get("action", "assign") == "assign" else 1,
                )
            with col_xd:
                default_xd = already.get("xdock", XDOCK_OPTIONS[0]) if action == "Asignar XDOCK" else XDOCK_OPTIONS[0]
                xd_sel = st.selectbox(
                    "XDOCK a asignar",
                    XDOCK_OPTIONS,
                    index=XDOCK_OPTIONS.index(default_xd) if default_xd in XDOCK_OPTIONS else 0,
                    key=f"xdsel_{key}",
                    disabled=(action == "Eliminar registros"),
                )
            with col_save:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 Guardar", key=f"save_{key}", use_container_width=True):
                    if action == "Asignar XDOCK" and xd_sel == XDOCK_OPTIONS[0]:
                        st.warning("Selecciona un XDOCK válido primero.")
                    else:
                        dec_entry = {
                            "action": "assign" if action == "Asignar XDOCK" else "delete",
                            "xdock":  xd_sel if action == "Asignar XDOCK" else "",
                            "guardado": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                        }
                        current_decisions[key] = dec_entry
                        save_decisions(current_decisions)
                        any_new = True
                        st.success(f"✅ Decisión guardada para **{id_s}**")

        if any_new:
            st.info("💡 Vuelve a procesar el archivo para aplicar las decisiones guardadas.")

    else:
        st.success("✅ Todos los registros tienen XDOCK asignado. No hay pendientes.")

    # ── Downloads ─────────────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">⬇️ Descargar Reportes</div>',
                unsafe_allow_html=True)
    fecha_str = datetime.date.today().strftime("%Y%m%d")

    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            label="📥 Descargar Reporte Completo Excel",
            data=excel_buf,
            file_name=f"GASO_REPORTE_{fecha_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d2:
        csv_buf = df_clean.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="📄 Descargar Base Limpia CSV",
            data=csv_buf,
            file_name=f"GASO_INOUT_LIMPIO_{fecha_str}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # Footer
    st.markdown(f"""
    <div style="text-align:center;padding:1.2rem;margin-top:1.5rem;
                border-top:1px solid #DCE8F5;color:#AAA;font-size:.75rem">
      GASO COMUNICACIONES · IN-OUT Report Processor v2.0 ·
      Generado el {datetime.date.today().strftime('%d/%m/%Y')}
    </div>
    """, unsafe_allow_html=True)
