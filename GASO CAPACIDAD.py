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
import base64
import tempfile
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Image as RLImage,
                                 Table, TableStyle, PageBreak, HRFlowable, KeepTogether)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY

# ─────────────────────────────────────────────────────────────────────────────
#  CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
DECISIONS_FILE = "gaso_decisions.json"   # persisted next to the script
LOGO_PATH = "GASO_COMUNICACIONES_LOGO.jpg"  # place beside the script

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


# Reglas para inferir tipo de pallet desde descripcion/clasificacion cuando
# el campo tipo_pallet está vacío. Orden importa: más específico primero.
DESC_PALLET_RULES = [
    # SOBREDIMENSIONADA
    ("antena gd",          "SOBREDIMENSIONADA"),
    ("antena grande",      "SOBREDIMENSIONADA"),
    ("antenas grandes",    "SOBREDIMENSIONADA"),
    ("galvanizado",        "SOBREDIMENSIONADA"),
    ("soporte",            "SOBREDIMENSIONADA"),   # cubre soportes, soporte desmontaje, etc.
    ("obra civil",         "SOBREDIMENSIONADA"),
    ("civil work",         "SOBREDIMENSIONADA"),
    ("base metalica",      "SOBREDIMENSIONADA"),
    ("base metálica",      "SOBREDIMENSIONADA"),
    ("vallen",             "SOBREDIMENSIONADA"),
    ("outdoor",            "SOBREDIMENSIONADA"),
    # EUROPALLET
    ("dsv",                "EUROPALLET"),          # pallet dsv / 1 pallet dsv / dsv-vallen
    ("hw sitio completo",  "EUROPALLET"),
    ("hw",                 "EUROPALLET"),          # HW solo o con tabuladores
    ("sde",                "EUROPALLET"),
    # ESTANDAR
    ("antena md",          "ESTANDAR"),
    ("antena mediana",     "ESTANDAR"),
    ("antena ch",          "EUROPALLET"),
    ("antena chica",       "EUROPALLET"),
    ("antena",             "ESTANDAR"),            # antena/antenas genérico → ESTANDAR
    ("gabinete",           "ESTANDAR"),
    ("complemento",        "ESTANDAR"),
    ("refurbish",          "ESTANDAR"),
    ("refaccion",          "ESTANDAR"),
    ("miscelaneo",         "ESTANDAR"),
    ("varios",             "ESTANDAR"),
    ("scrap",              "ESTANDAR"),
    ("logistica",          "ESTANDAR"),
    ("caja",               "ESTANDAR"),
    ("herraje",            "ESTANDAR"),
    ("sobrante",           "ESTANDAR"),
    ("cintillo",           "ESTANDAR"),
    ("bateria",            "ESTANDAR"),
    ("cable",              "ESTANDAR"),
    ("indoor",             "ESTANDAR"),
    ("rru",                "ESTANDAR"),
]


# Jerarquía de tamaño para elegir el tipo más grande cuando hay múltiples matches
TIPO_RANK = {"ESTANDAR": 1, "EUROPALLET": 2, "SOBREDIMENSIONADA": 3}


def _infer_tipo_from_desc(*fields):
    """
    Evalúa TODOS los segmentos del texto (separados por tab, coma, pipe, etc.)
    contra DESC_PALLET_RULES y devuelve el tipo MÁS GRANDE que encuentre.
    Si no hay ninguna coincidencia → "ESTANDAR" por default.
    """
    combined = " | ".join(norm(str(f)) for f in fields if f and str(f) not in ("nan", "none", ""))
    if not combined.strip():
        return "ESTANDAR"

    # Split on common separators so each segmento se evalúa por separado
    import re as _re
    segments = _re.split(r"[	,|/\n]+", combined)

    best = "ESTANDAR"
    for seg in segments:
        seg = seg.strip()
        if not seg:
            continue
        for keyword, tipo in DESC_PALLET_RULES:
            if keyword in seg:
                if TIPO_RANK.get(tipo, 0) > TIPO_RANK.get(best, 0):
                    best = tipo
                break  # una regla por segmento es suficiente
    return best


def normalize_tipo_pallet(df, cols):
    """
    1. Si tipo_pallet ya tiene valor reconocido → normalizar con PALLET_NORM.
    2. Si está vacío → inferir desde descripcion_material y clasificacion,
       tomando el tipo MÁS GRANDE entre todos los segmentos.
    3. Si nada coincide → ESTANDAR por default (nunca queda vacío).
    """
    desc_col  = cols.get("desc_material")
    clasi_col = None
    for c in df.columns:
        if "clasif" in c:
            clasi_col = c
            break

    resultado_inferencia = []

    def fix(row):
        val = row[cols["tipo_pallet"]]
        is_empty = pd.isna(val) or str(val).strip() in ("", "None", "nan")

        if not is_empty:
            n = norm(val)
            for k, v in PALLET_NORM.items():
                if k in n:
                    return v
            return "ESTANDAR"

        # Empty → infer (always returns a value, never None)
        desc  = row[desc_col]  if desc_col  and desc_col  in row.index else None
        clasi = row[clasi_col] if clasi_col and clasi_col in row.index else None
        inferred = _infer_tipo_from_desc(desc, clasi)

        resultado_inferencia.append({
            "ID_SITIO":      row.get(cols["id_sitio"], ""),
            "DESC":          str(desc)[:60],
            "CLASIF":        str(clasi)[:40],
            "TIPO_INFERIDO": inferred,
        })
        return inferred

    df[cols["tipo_pallet"]] = df.apply(fix, axis=1)

    global _last_tipo_inferences
    _last_tipo_inferences = resultado_inferencia

    return df


_last_tipo_inferences = []   # populated by normalize_tipo_pallet


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

    # Separar consolidados (NO. PALLET = 0) ANTES de fix_pallets
    # Son material pequeno encima de otra tarima: no impacta capacidad
    def _is_zero(val):
        try:
            return float(val) == 0
        except (TypeError, ValueError):
            return False

    mask_consol = df[cols["no_pallet"]].apply(_is_zero)
    df_consol   = df[mask_consol].copy()
    df          = df[~mask_consol].copy()
    logs.append(f"📌 Sitios consolidados (pallet=0): {len(df_consol)} | Para capacidad: {len(df):,}")

    df = fix_pallets(df, cols)
    df = normalize_tipo_pallet(df, cols)

    # Separar filas cuyo tipo_pallet no pudo inferirse → revisión manual de tipo
    n_inferred = len(_last_tipo_inferences)
    if n_inferred:
        logs.append(f"🔎 Tipo Pallet inferido por descripción: {n_inferred} filas")

    df = calc_m2(df, cols)
    df = assign_region(df, cols)

    # Enriquecer consolidados (sin M2 ni conteo de pallet)
    df_consol = normalize_tipo_pallet(df_consol, cols)
    df_consol["REGION"] = df_consol[cols["xdock"]].map(REGION_MAP).fillna("SIN REGION")
    df_consol["CIUDAD"] = df_consol[cols["xdock"]].map(CIUDAD_MAP).fillna("")
    df_consol["M2"]     = 0.0
    df["CONSOLIDADO"]       = False
    df_consol["CONSOLIDADO"] = True

    logs.append(f"📦 Pallets activos: {int(df[cols['no_pallet']].sum()):,} | M²: {df['M2'].sum():,.2f} | Consolidados: {len(df_consol)}")

    return df, df_consol, df_pending, logs, cols


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



# ─────────────────────────────────────────────────────────────────────────────
#  EXCEL EXPORT  (4 sheets: IN-OUT LIMPIO · REPORTE CLIENTE · RESUMEN EJECUTIVO · SITIOS CONSOLIDADOS)
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



def build_original_format_excel(df_clean, df_consol, cols):
    """Reproduce the original IN-OUT workbook header style with the clean data."""
    wb   = openpyxl.Workbook()
    ws   = wb.active
    ws.title = "IN-OUT"
    fecha = datetime.date.today().strftime("%d/%m/%Y")

    # ── Colour palette matching the original ──────────────────────────────────
    C_DARK_BLUE = "002060"   # Row 1 / headers INGRESO
    C_DARK_GRN  = "0B6A0B"   # headers SALIDA
    C_GOLD      = "FFD966"   # headers INVENTARIO
    C_LT_BLUE   = "D9E1F2"   # sub-header INGRESO
    C_LT_GRN    = "C6EFCE"   # sub-header SALIDA
    C_LT_GOLD   = "FFF2CC"   # sub-header INVENTARIO
    C_WHITE     = "FFFFFF"

    def _mk_fill(hex_col):
        return PatternFill("solid", fgColor=hex_col)

    def _mk_font(bold=False, color=C_WHITE, size=11):
        return Font(name="Calibri", bold=bold, color=color, size=size)

    def _ctr():
        return Alignment(horizontal="center", vertical="center", wrap_text=True)

    def _lft():
        return Alignment(horizontal="left", vertical="center")

    # Column layout: map to original widths (A=CARRIER … )
    # We output only the active/relevant columns
    out_cols_names = [
        "CARRIER", "XDOCK", "FECHA DE INGRESO", "HORA DE INGRESO",
        "FOLIO ALMACEN ORIGEN", "ESTATUS", "CLASIFICACION DE MATERIAL",
        "TIPO DE MATERIAL", "ID SITIO", "NOMBRE DE SITIO",
        "FOLIO CLIENTE ( PDM )", "ID PALLET", "NO. DE PALLET",
        "TIPO DE PALLET", "FOLIO WMS GASO ( UNICO X PALLET )",
        "DESCRIPCION MATERIAL", "PROYECTO", "SUB-PROYECTO",
        "FOLIO ENRUTADO CLIENTE", "ESTATUS SALIDA",
        "REGION", "CIUDAD", "M2", "CONSOLIDADO",
    ]
    # Map to df columns (normalized)
    col_lookup = {
        "CARRIER":                     cols["carrier"],
        "XDOCK":                       cols["xdock"],
        "FECHA DE INGRESO":            resolve_col(df_clean, ["fecha de ingreso"]),
        "HORA DE INGRESO":             resolve_col(df_clean, ["hora de ingreso"]),
        "FOLIO ALMACEN ORIGEN":        cols["folio"],
        "ESTATUS":                     cols["estatus"],
        "CLASIFICACION DE MATERIAL":   resolve_col(df_clean, ["clasificacion de material"]),
        "TIPO DE MATERIAL":            cols["tipo_material"],
        "ID SITIO":                    cols["id_sitio"],
        "NOMBRE DE SITIO":             cols["nombre_sitio"],
        "FOLIO CLIENTE ( PDM )":       resolve_col(df_clean, ["folio cliente ( pdm )"]),
        "ID PALLET":                   resolve_col(df_clean, ["id pallet"]),
        "NO. DE PALLET":               cols["no_pallet"],
        "TIPO DE PALLET":              cols["tipo_pallet"],
        "FOLIO WMS GASO ( UNICO X PALLET )": resolve_col(df_clean, ["folio wms gaso ( unico x pallet )"]),
        "DESCRIPCION MATERIAL":        cols.get("desc_material"),
        "PROYECTO":                    resolve_col(df_clean, ["proyecto"]),
        "SUB-PROYECTO":                resolve_col(df_clean, ["sub-proyecto"]),
        "FOLIO ENRUTADO CLIENTE":      resolve_col(df_clean, ["folio enrutado cliente"]),
        "ESTATUS SALIDA":              cols["est_salida"],
        "REGION":                      "REGION",
        "CIUDAD":                      "CIUDAD",
        "M2":                          "M2",
        "CONSOLIDADO":                 "CONSOLIDADO",
    }
    # Original column widths (approx)
    orig_widths = {
        "CARRIER": 16.8, "XDOCK": 14.8, "FECHA DE INGRESO": 16.8,
        "HORA DE INGRESO": 12.8, "FOLIO ALMACEN ORIGEN": 22.8,
        "ESTATUS": 14.8, "CLASIFICACION DE MATERIAL": 26.8,
        "TIPO DE MATERIAL": 20.8, "ID SITIO": 18.8, "NOMBRE DE SITIO": 28.8,
        "FOLIO CLIENTE ( PDM )": 20.8, "ID PALLET": 16.8,
        "NO. DE PALLET": 16.8, "TIPO DE PALLET": 16.8,
        "FOLIO WMS GASO ( UNICO X PALLET )": 28.8,
        "DESCRIPCION MATERIAL": 32.8, "PROYECTO": 18.8, "SUB-PROYECTO": 18.8,
        "FOLIO ENRUTADO CLIENTE": 22.8, "ESTATUS SALIDA": 16.8,
        "REGION": 16.8, "CIUDAD": 14.8, "M2": 10.0, "CONSOLIDADO": 12.0,
    }

    n_cols = len(out_cols_names)
    last_col_letter = get_column_letter(n_cols)

    # Row 1: REPORTE IN-OUT LIMPIO (dark blue, merged)
    ws.merge_cells(f"A1:{last_col_letter}1")
    c = ws["A1"]
    c.value = "REPORTE IN-OUT"
    c.font  = _mk_font(bold=True, color=C_WHITE, size=13)
    c.fill  = _mk_fill(C_DARK_BLUE)
    c.alignment = _ctr()
    ws.row_dimensions[1].height = 22

    # Row 2: GASO COMUNICACIONES
    ws.merge_cells(f"A2:{last_col_letter}2")
    c = ws["A2"]
    c.value = f"GASO COMUNICACIONES  –  Base Limpia Procesada  |  {fecha}"
    c.font  = _mk_font(bold=True, color=C_DARK_BLUE, size=11)
    c.fill  = _mk_fill("F4F6F9")
    c.alignment = _ctr()
    ws.row_dimensions[2].height = 18

    # Row 3: Section labels INGRESO / SALIDA / INVENTARIO
    # INGRESO spans cols 1-20, SALIDA none (no output), REGION cols 21+
    ingreso_end   = min(20, n_cols)
    region_start  = 21
    ws.merge_cells(f"A3:{get_column_letter(ingreso_end)}3")
    c = ws["A3"]; c.value = "INGRESO"
    c.font = _mk_font(bold=True); c.fill = _mk_fill(C_DARK_BLUE); c.alignment = _ctr()
    if region_start <= n_cols:
        ws.merge_cells(f"{get_column_letter(region_start)}3:{last_col_letter}3")
        c = ws[f"{get_column_letter(region_start)}3"]; c.value = "INVENTARIO / MÉTRICAS"
        c.font = _mk_font(bold=True, color="333333"); c.fill = _mk_fill(C_GOLD); c.alignment = _ctr()
    ws.row_dimensions[3].height = 18

    # Row 4: Sub-headers INPUT / INVENTARIO
    ws.merge_cells(f"A4:{get_column_letter(ingreso_end)}4")
    c = ws["A4"]; c.value = "INPUT"
    c.font = _mk_font(bold=True, color=C_DARK_BLUE); c.fill = _mk_fill(C_LT_BLUE); c.alignment = _ctr()
    if region_start <= n_cols:
        ws.merge_cells(f"{get_column_letter(region_start)}4:{last_col_letter}4")
        c = ws[f"{get_column_letter(region_start)}4"]; c.value = "INVENTARIO"
        c.font = _mk_font(bold=True, color="333333"); c.fill = _mk_fill(C_LT_GOLD); c.alignment = _ctr()
    ws.row_dimensions[4].height = 18

    # Row 5: Column headers
    ws.row_dimensions[5].height = 30
    for ci, col_name in enumerate(out_cols_names, 1):
        is_region_col = ci >= region_start
        c = ws.cell(row=5, column=ci, value=col_name)
        c.font  = _mk_font(bold=True, color=C_WHITE, size=11)
        c.fill  = _mk_fill(C_GOLD if is_region_col else C_DARK_BLUE)
        if is_region_col:
            c.font = _mk_font(bold=True, color="333333", size=11)
        c.alignment = _ctr()
        c.border = BORDER
        ws.column_dimensions[get_column_letter(ci)].width = orig_widths.get(col_name, 14.0)

    # Data rows - combine clean + consolidados
    df_all = pd.concat([df_clean, df_consol], ignore_index=True)
    thin = Side(style="thin", color="CCCCCC")
    data_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ri, (_, row) in enumerate(df_all.iterrows()):
        er = 6 + ri
        alt = ri % 2 == 1
        base_fill = PatternFill("solid", fgColor="EBF0F7" if alt else "FFFFFF")
        for ci, col_name in enumerate(out_cols_names, 1):
            df_col = col_lookup.get(col_name)
            val = row[df_col] if df_col and df_col in row.index else None
            c = ws.cell(row=er, column=ci, value=val)
            c.font      = Font(name="Calibri", size=10)
            c.alignment = _lft()
            c.border    = data_border
            c.fill      = base_fill

    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A5:{last_col_letter}5"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

def build_excel_output(df_clean, df_consol, cols):
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

    # ── 3. RESUMEN EJECUTIVO ─────────────────────────────────────────────────
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


    # ── 5. SITIOS CONSOLIDADOS ───────────────────────────────────────────────
    ws5 = wb.create_sheet("SITIOS CONSOLIDADOS")
    _title_block(ws5, "SITIOS CONSOLIDADOS",
                 "Material consolidado sobre otra tarima – no impacta en capacidad ni M²", fecha,
                 n_cols=min(20, len(df_consol.columns) if len(df_consol) > 0 else 8))
    if len(df_consol) > 0:
        # Drop helper column before export
        consol_export = df_consol.drop(columns=["CONSOLIDADO"], errors="ignore").reset_index(drop=True)
        _write_table(ws5, consol_export, start_row=5)
        # Summary by XDOCK
        sum_row = 5 + len(consol_export) + 3
        ws5.cell(row=sum_row, column=1, value="RESUMEN POR CROSSDOCK").font =             Font(name="Calibri", bold=True, size=10, color=HEX_BLUE)
        sum_hdr = ["CIUDAD", "XDOCK", "CARRIER", "SITIOS CONSOLIDADOS"]
        for ci, h in enumerate(sum_hdr, 1):
            c = ws5.cell(row=sum_row + 1, column=ci, value=h)
            c.font = HDR_FONT; c.fill = HDR_FILL; c.border = BORDER; c.alignment = CTR
            ws5.column_dimensions[get_column_letter(ci)].width = 22
        grp_c = (df_consol.groupby([cols["xdock"], cols["carrier"]])
                 .size().reset_index(name="SITIOS"))
        for ri2, row2 in grp_c.iterrows():
            er2 = sum_row + 2 + ri2
            vals = [CIUDAD_MAP.get(row2[cols["xdock"]], row2[cols["xdock"]]),
                    row2[cols["xdock"]], row2[cols["carrier"]], row2["SITIOS"]]
            for ci, v in enumerate(vals, 1):
                c = ws5.cell(row=er2, column=ci, value=v)
                c.fill = _alt_fill(ri2); c.border = BORDER
                c.font = DATA_FONT; c.alignment = CTR
    else:
        ws5["A5"] = "No hay sitios consolidados en esta carga."
        ws5["A5"].font = Font(name="Calibri", size=10, color="555555")


    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────────────────────
#  PLOTLY CHARTS
# ─────────────────────────────────────────────────────────────────────────────

def _smart_analysis(xd, pct, total_pal, m2_ocp, cap, tipo_dist):
    """Generate an intelligent 2-3 sentence analysis for a crossdock."""
    ciudad = CIUDAD_MAP.get(xd, xd)
    pct100 = round(pct * 100, 1)

    if pct > 1.0:
        main = (f"{ciudad} opera al {pct100}% de su capacidad, superando el límite de {cap} m². "
                f"Con {total_pal} pallets y {m2_ocp:.0f} m² ocupados se requiere acción inmediata: "
                f"revisar si existen capturas duplicadas o coordinar salidas urgentes.")
    elif pct > 0.90:
        main = (f"{ciudad} registra una ocupación crítica del {pct100}% ({m2_ocp:.0f}/{cap} m²). "
                f"Queda menos del 10% de capacidad disponible — se recomienda programar salidas "
                f"antes de recibir nuevas entradas.")
    elif pct > 0.70:
        main = (f"{ciudad} está en zona de alerta con {pct100}% de ocupación ({m2_ocp:.0f}/{cap} m²). "
                f"El crossdock puede absorber entregas moderadas pero debe monitorearse semanalmente.")
    elif pct > 0.30:
        main = (f"{ciudad} opera a un nivel saludable del {pct100}% ({m2_ocp:.0f}/{cap} m²), "
                f"con {round(cap-m2_ocp,0):.0f} m² disponibles para nuevas entradas.")
    else:
        if total_pal == 0:
            main = (f"{ciudad} no registra inventario activo en este periodo. "
                    f"Verificar si el crossdock está operativo o si existen capturas pendientes.")
        else:
            main = (f"{ciudad} presenta ocupación baja del {pct100}% ({total_pal} pallets). "
                    f"Capacidad ampliamente disponible ({round(cap-m2_ocp,0):.0f} m² libres).")

    # Tipo de material note
    if tipo_dist:
        top_tipo = max(tipo_dist, key=tipo_dist.get)
        main += f" El tipo de pallet predominante es {top_tipo}."

    return main


def generate_pdf(df_clean, df_consol, cols, region_filter="Todas"):
    """Build an executive PDF report with charts and smart analysis per crossdock."""
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
    from reportlab.platypus import Table, TableStyle, PageBreak, HRFlowable, KeepTogether
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
    from reportlab.lib.pagesizes import A4, landscape

    # Filter by region if needed
    if region_filter != "Todas":
        df_plot = df_clean[df_clean["REGION"] == region_filter].copy()
        xdocks_plot = [xd for xd in CAPACIDADES if REGION_MAP.get(xd) == region_filter]
    else:
        df_plot = df_clean.copy()
        xdocks_plot = list(CAPACIDADES.keys())

    xdock_col  = cols["xdock"]
    carrier_col = cols["carrier"]
    mat_col    = cols["tipo_material"]
    pallet_col = cols["no_pallet"]
    fecha_str  = datetime.date.today().strftime("%d de %B de %Y")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2.5*cm, bottomMargin=2*cm,
        title="Reporte Ejecutivo Gaso Comunicaciones",
        author="Gaso Comunicaciones",
    )

    # ── Styles ────────────────────────────────────────────────────────────────
    RL_BLUE  = colors.HexColor("#1A3A6B")
    RL_LIGHT = colors.HexColor("#2E6DB4")
    RL_ACCENT= colors.HexColor("#4A90D9")
    RL_GREEN = colors.HexColor("#1E8449")
    RL_AMBER = colors.HexColor("#E67E22")
    RL_RED   = colors.HexColor("#C0392B")
    RL_DRED  = colors.HexColor("#7B241C")
    RL_LGRAY = colors.HexColor("#F4F6F9")
    RL_WHITE = colors.white

    styles = getSampleStyleSheet()

    def _sty(name, **kw):
        return ParagraphStyle(name, **kw)

    S_COVER_TITLE = _sty("CoverTitle", fontSize=28, textColor=RL_WHITE,
                          fontName="Helvetica-Bold", alignment=TA_CENTER, spaceAfter=6)
    S_COVER_SUB   = _sty("CoverSub",   fontSize=13, textColor=colors.HexColor("#C8D8EC"),
                          fontName="Helvetica", alignment=TA_CENTER, spaceAfter=4)
    S_COVER_DATE  = _sty("CoverDate",  fontSize=10, textColor=colors.HexColor("#AABBCC"),
                          fontName="Helvetica", alignment=TA_CENTER)
    S_H1    = _sty("H1",   fontSize=14, textColor=RL_BLUE, fontName="Helvetica-Bold",
                   spaceBefore=14, spaceAfter=4, borderPad=0)
    S_H2    = _sty("H2",   fontSize=11, textColor=RL_LIGHT, fontName="Helvetica-Bold",
                   spaceBefore=8, spaceAfter=3)
    S_BODY  = _sty("Body", fontSize=9,  textColor=colors.HexColor("#333333"),
                   fontName="Helvetica", leading=14, alignment=TA_JUSTIFY, spaceAfter=6)
    S_SMALL = _sty("Small",fontSize=8,  textColor=colors.HexColor("#777777"),
                   fontName="Helvetica", leading=11)
    S_CTR   = _sty("Ctr",  fontSize=9,  textColor=colors.HexColor("#333333"),
                   fontName="Helvetica", alignment=TA_CENTER)
    S_BOLD  = _sty("Bold", fontSize=9,  textColor=RL_BLUE, fontName="Helvetica-Bold",
                   leading=13)

    def _hr():
        return HRFlowable(width="100%", thickness=1, color=RL_ACCENT, spaceAfter=6, spaceBefore=2)

    def _spacer(h=0.3):
        return Spacer(1, h*cm)

    def _fig_to_image(fig, width_cm=16, height_cm=8):
        """Export a plotly figure to a ReportLab Image flowable."""
        img_bytes = fig.to_image(format="png", width=int(width_cm/16*1200),
                                  height=int(height_cm/8*600), scale=2)
        return RLImage(io.BytesIO(img_bytes), width=width_cm*cm, height=height_cm*cm)

    story = []

    # ══════════════════════════════════════════════════════════════════════════
    # COVER PAGE
    # ══════════════════════════════════════════════════════════════════════════
    # Blue cover rectangle via a 1-cell table
    region_label = region_filter if region_filter != "Todas" else "Todas las Regiones"
    cover_data = [[Paragraph("GASO COMUNICACIONES", S_COVER_TITLE)],
                  [Paragraph("Reporte Ejecutivo de Ocupación de Inventario", S_COVER_SUB)],
                  [Paragraph(f"Región: {region_label}", S_COVER_SUB)],
                  [Paragraph(fecha_str, S_COVER_DATE)]]
    cover_tbl = Table(cover_data, colWidths=[17*cm])
    cover_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), RL_BLUE),
        ("TOPPADDING",    (0,0), (-1,0),  60),
        ("BOTTOMPADDING", (0,-1),(-1,-1), 60),
        ("LEFTPADDING",   (0,0), (-1,-1), 20),
        ("RIGHTPADDING",  (0,0), (-1,-1), 20),
        ("ROWBACKGROUNDS",(0,0), (-1,-1), [RL_BLUE]),
    ]))

    # Logo if available
    if os.path.exists(LOGO_PATH):
        try:
            logo = RLImage(LOGO_PATH, width=4*cm, height=2*cm)
            logo.hAlign = "CENTER"
            story.append(_spacer(1))
            story.append(logo)
            story.append(_spacer(0.5))
        except Exception:
            story.append(_spacer(3))
    else:
        story.append(_spacer(5))

    story.append(cover_tbl)
    story.append(_spacer(1))

    # KPI summary boxes on cover
    total_pal  = int(df_plot[pallet_col].sum())
    total_m2   = round(df_plot["M2"].sum(), 2)
    cap_region = sum(CAPACIDADES[xd] for xd in xdocks_plot)
    pct_global = round(total_m2 / cap_region * 100, 1) if cap_region > 0 else 0
    disponible = round(cap_region - total_m2, 2)
    n_consol   = len(df_consol[df_consol[xdock_col].isin(xdocks_plot)] if len(df_consol) > 0 else df_consol)

    def _pct_color(p):
        if p > 100: return RL_DRED
        if p > 90:  return RL_RED
        if p > 70:  return RL_AMBER
        return RL_GREEN

    kpi_rows = [[
        Paragraph(f"<b>{total_pal:,}</b><br/>Pallets Activos", S_CTR),
        Paragraph(f"<b>{total_m2:,.0f} m²</b><br/>M² Ocupados", S_CTR),
        Paragraph(f"<b>{cap_region:,} m²</b><br/>Capacidad Total", S_CTR),
        Paragraph(f"<b>{pct_global}%</b><br/>% Ocupación", S_CTR),
        Paragraph(f"<b>{disponible:,.0f} m²</b><br/>Disponible", S_CTR),
    ]]
    kpi_tbl = Table(kpi_rows, colWidths=[3.3*cm]*5)
    pct_bg = _pct_color(pct_global)
    kpi_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(0,0), colors.HexColor("#EBF0F7")),
        ("BACKGROUND", (1,0),(1,0), colors.HexColor("#EBF0F7")),
        ("BACKGROUND", (2,0),(2,0), colors.HexColor("#EBF0F7")),
        ("BACKGROUND", (3,0),(3,0), pct_bg),
        ("BACKGROUND", (4,0),(4,0), colors.HexColor("#EBF0F7")),
        ("TEXTCOLOR",  (3,0),(3,0), RL_WHITE),
        ("FONTNAME",   (3,0),(3,0), "Helvetica-Bold"),
        ("ALIGN",      (0,0),(-1,-1), "CENTER"),
        ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
        ("ROWHEIGHT",  (0,0),(-1,-1), 1.5*cm),
        ("BOX",        (0,0),(-1,-1), 0.5, RL_ACCENT),
        ("INNERGRID",  (0,0),(-1,-1), 0.3, colors.HexColor("#CCDDEE")),
        ("TOPPADDING", (0,0),(-1,-1), 8),
        ("BOTTOMPADDING",(0,0),(-1,-1),8),
    ]))
    story.append(kpi_tbl)
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 2: RESUMEN DE OCUPACIÓN
    # ══════════════════════════════════════════════════════════════════════════
    story.append(Paragraph("1. Resumen de Ocupación por Crossdock", S_H1))
    story.append(_hr())

    # Build occupancy fig
    ocp = []
    for xd in xdocks_plot:
        cap    = CAPACIDADES[xd]
        m2_ocp = round(df_plot[df_plot[xdock_col]==xd]["M2"].sum(), 2)
        pct    = m2_ocp / cap if cap > 0 else 0
        ocp.append({"Ciudad": CIUDAD_MAP.get(xd,xd), "pct": round(pct*100,1),
                    "M2_ocp": m2_ocp, "cap": cap, "disp": round(cap-m2_ocp,1), "xd": xd})
    df_ocp = pd.DataFrame(ocp).sort_values("pct", ascending=True)

    fig_ocp = go.Figure(go.Bar(
        x=df_ocp["pct"], y=df_ocp["Ciudad"], orientation="h",
        marker_color=[_color_pct(p/100) for p in df_ocp["pct"]],
        text=[f"{p}%" for p in df_ocp["pct"]], textposition="outside",
    ))
    for v, lbl, clr in [(70,"70%","#E67E22"),(90,"90%","#C0392B"),(100,"100%","#7B241C")]:
        fig_ocp.add_vline(x=v, line_dash="dot", line_color=clr,
                          annotation_text=lbl, annotation_font_color=clr)
    fig_ocp.update_layout(
        title="", xaxis_title="% Ocupación", yaxis_title="",
        plot_bgcolor="white", paper_bgcolor="white",
        height=350, margin=dict(l=10,r=80,t=10,b=30),
        font=dict(family="Helvetica", color="#1A3A6B"),
        xaxis=dict(range=[0, max(df_ocp["pct"].max()*1.18, 115)]),
    )
    story.append(_fig_to_image(fig_ocp, 16, 7.5))
    story.append(_spacer(0.3))

    # ── Per-crossdock analysis table ──────────────────────────────────────────
    story.append(Paragraph("Análisis por Crossdock", S_H2))
    tbl_hdr = ["Ciudad", "Cap. m²", "M² Ocup.", "% Ocup.", "Disponible", "Status"]
    tbl_rows = [tbl_hdr]
    for row_d in ocp:
        xd    = row_d["xd"]
        pct_v = row_d["pct"]
        st_lbl = ("SATURADO" if pct_v>100 else "CRÍTICO" if pct_v>90
                  else "ALERTA" if pct_v>70 else "NORMAL")
        tbl_rows.append([
            row_d["Ciudad"],
            f"{row_d['cap']:,}",
            f"{row_d['M2_ocp']:,.0f}",
            f"{pct_v:.1f}%",
            f"{row_d['disp']:,.0f}",
            st_lbl,
        ])
    ocp_tbl = Table(tbl_rows, colWidths=[3.2*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.4*cm, 2.4*cm])

    def _status_color(s):
        if "SATURADO" in s: return RL_DRED
        if "CRÍTICO"  in s: return RL_RED
        if "ALERTA"   in s: return RL_AMBER
        return RL_GREEN

    ts = [
        ("BACKGROUND",   (0,0),(-1,0), RL_BLUE),
        ("TEXTCOLOR",    (0,0),(-1,0), RL_WHITE),
        ("FONTNAME",     (0,0),(-1,0), "Helvetica-Bold"),
        ("FONTSIZE",     (0,0),(-1,-1), 8),
        ("ALIGN",        (0,0),(-1,-1), "CENTER"),
        ("VALIGN",       (0,0),(-1,-1), "MIDDLE"),
        ("ROWHEIGHT",    (0,0),(-1,-1), 0.6*cm),
        ("BOX",          (0,0),(-1,-1), 0.5, RL_ACCENT),
        ("INNERGRID",    (0,0),(-1,-1), 0.3, colors.HexColor("#CCDDEE")),
        ("TOPPADDING",   (0,0),(-1,-1), 4),
        ("BOTTOMPADDING",(0,0),(-1,-1), 4),
    ]
    for ri, row_d in enumerate(ocp[: ], 1):
        bg = colors.HexColor("#EBF0F7") if ri%2==0 else colors.white
        ts.append(("BACKGROUND", (0,ri),(-2,ri), bg))
        pct_v = row_d["pct"]
        ts.append(("BACKGROUND", (-1,ri),(-1,ri), _status_color(
            "SATURADO" if pct_v>100 else "CRÍTICO" if pct_v>90 else "ALERTA" if pct_v>70 else "NORMAL")))
        ts.append(("TEXTCOLOR",  (-1,ri),(-1,ri), RL_WHITE))
        ts.append(("FONTNAME",   (-1,ri),(-1,ri), "Helvetica-Bold"))

    ocp_tbl.setStyle(TableStyle(ts))
    story.append(ocp_tbl)
    story.append(_spacer(0.4))

    # Global analysis paragraph
    sat = [r for r in ocp if r["pct"]>100]
    crit= [r for r in ocp if 90<r["pct"]<=100]
    ok  = [r for r in ocp if r["pct"]<=70]
    analysis_lines = []
    if sat:
        analysis_lines.append(
            f"<b>Saturación detectada:</b> {', '.join(r['Ciudad'] for r in sat)} superan el 100% de capacidad. "
            "Esto puede indicar capturas sin confirmación de salida o errores de registro — se recomienda auditoria inmediata."
        )
    if crit:
        analysis_lines.append(
            f"<b>Nivel crítico:</b> {', '.join(r['Ciudad'] for r in crit)} están entre el 90-100%. "
            "Priorizar coordinar salidas antes de nuevas recepciones."
        )
    if not sat and not crit:
        analysis_lines.append("El inventario global se encuentra en niveles manejables. Continuar monitoreo semanal.")
    if ok:
        analysis_lines.append(
            f"{', '.join(r['Ciudad'] for r in ok)} tienen capacidad ampliamente disponible."
        )
    for line in analysis_lines:
        story.append(Paragraph(line, S_BODY))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 3: M² y MATERIAL
    # ══════════════════════════════════════════════════════════════════════════
    story.append(Paragraph("2. Análisis de M² y Tipo de Material", S_H1))
    story.append(_hr())

    # Stacked M2 chart
    fig_m2 = go.Figure([
        go.Bar(x=df_ocp["Ciudad"], y=df_ocp["M2_ocp"], name="M² Ocupados",
               marker_color=GASO_LIGHT),
        go.Bar(x=df_ocp["Ciudad"], y=df_ocp["disp"], name="Disponible",
               marker_color="#D5E8F5"),
    ])
    fig_m2.update_layout(
        barmode="stack", title="", xaxis_title="", yaxis_title="m²",
        plot_bgcolor="white", paper_bgcolor="white",
        height=300, margin=dict(l=10,r=10,t=10,b=60),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        font=dict(family="Helvetica", color="#1A3A6B"),
    )
    story.append(_fig_to_image(fig_m2, 16, 6.5))
    story.append(Paragraph(
        f"La capacidad total de la región seleccionada es de <b>{cap_region:,} m²</b>. "
        f"Actualmente se ocupan <b>{total_m2:,.0f} m²</b>, dejando <b>{disponible:,.0f} m²</b> disponibles "
        f"({round(100-pct_global,1)}% de holgura). "
        f"El factor de pasillos del 20% ya está incluido en todos los cálculos.",
        S_BODY))
    story.append(_spacer(0.3))

    # Material pie + heatmap side by side
    mat_grp = df_plot.groupby(mat_col)[pallet_col].sum().reset_index()
    mat_grp.columns = ["Tipo", "Pallets"]
    fig_pie = px.pie(mat_grp, names="Tipo", values="Pallets", hole=0.42,
                     color_discrete_sequence=[GASO_BLUE, GASO_ACCENT, "#E67E22", "#8E44AD", "#1E8449"])
    fig_pie.update_traces(textinfo="percent+label")
    fig_pie.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                          showlegend=False, height=280,
                          margin=dict(l=10,r=10,t=10,b=10),
                          font=dict(family="Helvetica", color="#1A3A6B"))

    heat = df_plot.groupby([xdock_col, mat_col])["M2"].sum().reset_index()
    hp   = heat.pivot(index=xdock_col, columns=mat_col, values="M2").fillna(0)
    hp.index = [CIUDAD_MAP.get(x,x) for x in hp.index]
    fig_heat = px.imshow(hp.round(1), text_auto=".0f", aspect="auto",
                         color_continuous_scale=[[0,"#EBF5FB"],[0.5,GASO_ACCENT],[1,GASO_BLUE]])
    fig_heat.update_layout(plot_bgcolor="white", paper_bgcolor="white", height=280,
                            margin=dict(l=10,r=10,t=30,b=10),
                            font=dict(family="Helvetica", color="#1A3A6B"),
                            coloraxis_showscale=False)

    story.append(Paragraph("Distribución por Tipo de Material", S_H2))
    side_data = [[_fig_to_image(fig_pie, 7.8, 6), _fig_to_image(fig_heat, 8.8, 6)]]
    side_tbl  = Table(side_data, colWidths=[8*cm, 9*cm])
    side_tbl.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),
                                   ("LEFTPADDING",(0,0),(-1,-1),0),
                                   ("RIGHTPADDING",(0,0),(-1,-1),4)]))
    story.append(side_tbl)

    top_mat = mat_grp.sort_values("Pallets", ascending=False).iloc[0]["Tipo"] if len(mat_grp) else "N/A"
    story.append(Paragraph(
        f"El tipo de material predominante es <b>{top_mat}</b>. "
        "El mapa de calor muestra la distribución de m² por tipo en cada crossdock, "
        "permitiendo identificar concentraciones de material específico.",
        S_BODY))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 4: CARRIER y ANÁLISIS INDIVIDUAL
    # ══════════════════════════════════════════════════════════════════════════
    story.append(Paragraph("3. Análisis por Carrier y Crossdock", S_H1))
    story.append(_hr())

    car_grp = df_plot.groupby(carrier_col)[pallet_col].sum().reset_index()
    car_grp.columns = ["Carrier", "Pallets"]
    fig_car = px.bar(car_grp.sort_values("Pallets", ascending=False),
                     x="Carrier", y="Pallets", text="Pallets",
                     color="Carrier",
                     color_discrete_sequence=[GASO_BLUE, GASO_ACCENT, "#E67E22", "#8E44AD"])
    fig_car.update_traces(textposition="outside")
    fig_car.update_layout(showlegend=False, plot_bgcolor="white", paper_bgcolor="white",
                          height=300, margin=dict(l=10,r=10,t=10,b=40),
                          font=dict(family="Helvetica", color="#1A3A6B"),
                          xaxis_title="", yaxis_title="Pallets")
    story.append(_fig_to_image(fig_car, 16, 6.5))

    top_carrier = car_grp.sort_values("Pallets", ascending=False).iloc[0]["Carrier"] if len(car_grp) else "N/A"
    top_car_pal = int(car_grp.sort_values("Pallets", ascending=False).iloc[0]["Pallets"]) if len(car_grp) else 0
    story.append(Paragraph(
        f"El carrier con mayor volumen de inventario es <b>{top_carrier}</b> con <b>{top_car_pal:,} pallets</b>. "
        "Un alto volumen de un solo carrier puede representar un riesgo operativo "
        "si existen retrasos en sus salidas programadas.",
        S_BODY))
    story.append(_spacer(0.4))

    # ── Individual crossdock analysis ─────────────────────────────────────────
    story.append(Paragraph("4. Análisis Individual por Crossdock", S_H1))
    story.append(_hr())

    for row_d in sorted(ocp, key=lambda r: r["pct"], reverse=True):
        xd     = row_d["xd"]
        ciudad = row_d["Ciudad"]
        cap    = row_d["cap"]
        m2_ocp = row_d["M2_ocp"]
        pct_v  = row_d["pct"]
        pal    = int(df_plot[df_plot[xdock_col]==xd][pallet_col].sum())

        tipo_dist = df_plot[df_plot[xdock_col]==xd].groupby(cols["tipo_pallet"])[pallet_col].sum().to_dict()

        analysis_text = _smart_analysis(xd, pct_v/100, pal, m2_ocp, cap, tipo_dist)
        status_lbl    = ("🔴 SATURADO" if pct_v>100 else "🔴 CRÍTICO" if pct_v>90
                         else "🟡 ALERTA" if pct_v>70 else "🟢 NORMAL")
        st_color = (_pct_color(pct_v/100))

        block = [
            Paragraph(f"{ciudad}  —  {status_lbl}", S_H2),
            Table([[
                Paragraph(f"<b>Pallets:</b> {pal:,}", S_BOLD),
                Paragraph(f"<b>M² Ocup.:</b> {m2_ocp:,.0f}", S_BOLD),
                Paragraph(f"<b>Capacidad:</b> {cap:,} m²", S_BOLD),
                Paragraph(f"<b>Ocupación:</b> {pct_v:.1f}%", S_BOLD),
            ]], colWidths=[3.8*cm]*4,
               style=TableStyle([
                   ("BACKGROUND",(0,0),(-1,-1), colors.HexColor("#F4F6F9")),
                   ("BOX",(0,0),(-1,-1),0.5,RL_ACCENT),
                   ("INNERGRID",(0,0),(-1,-1),0.2,colors.HexColor("#CCDDEE")),
                   ("ALIGN",(0,0),(-1,-1),"CENTER"),
                   ("TOPPADDING",(0,0),(-1,-1),5),
                   ("BOTTOMPADDING",(0,0),(-1,-1),5),
               ])),
            _spacer(0.2),
            Paragraph(analysis_text, S_BODY),
            _spacer(0.2),
            _hr(),
        ]
        story.append(KeepTogether(block))

    # ══════════════════════════════════════════════════════════════════════════
    # LAST PAGE: SITIOS CONSOLIDADOS + FOOTER
    # ══════════════════════════════════════════════════════════════════════════
    story.append(PageBreak())
    story.append(Paragraph("5. Sitios Consolidados", S_H1))
    story.append(_hr())

    df_cons_region = df_consol[df_consol[xdock_col].isin(xdocks_plot)] if len(df_consol) > 0 else df_consol
    n_consol_r = len(df_cons_region)
    story.append(Paragraph(
        f"Se identificaron <b>{n_consol_r} sitios consolidados</b> (pallets con valor 0) "
        "que corresponden a material pequeño ubicado sobre otra tarima. "
        "Estos registros <b>no impactan el cálculo de capacidad ni de m²</b> "
        "pero se mantienen en el inventario para trazabilidad.",
        S_BODY))

    if n_consol_r > 0:
        grp_c = df_cons_region.groupby([xdock_col, carrier_col]).size().reset_index(name="Sitios")
        grp_c["Ciudad"] = grp_c[xdock_col].map(CIUDAD_MAP)
        cons_data  = [["Ciudad", "Carrier", "Sitios Consolidados"]]
        for _, r in grp_c.iterrows():
            cons_data.append([r["Ciudad"], r[carrier_col], str(r["Sitios"])])
        cons_tbl = Table(cons_data, colWidths=[5*cm, 5*cm, 5*cm])
        cons_tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0), RL_BLUE),
            ("TEXTCOLOR", (0,0),(-1,0), RL_WHITE),
            ("FONTNAME",  (0,0),(-1,0), "Helvetica-Bold"),
            ("FONTSIZE",  (0,0),(-1,-1), 9),
            ("ALIGN",     (0,0),(-1,-1), "CENTER"),
            ("ROWHEIGHT", (0,0),(-1,-1), 0.65*cm),
            ("BOX",       (0,0),(-1,-1), 0.5, RL_ACCENT),
            ("INNERGRID", (0,0),(-1,-1), 0.3, colors.HexColor("#CCDDEE")),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#EBF0F7")]),
        ]))
        story.append(cons_tbl)

    # Footer note
    story.append(_spacer(1))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CCDDEE")))
    story.append(Paragraph(
        f"Reporte generado automáticamente por el sistema IN-OUT de Gaso Comunicaciones  |  {fecha_str}  |  Confidencial",
        _sty("Footer", fontSize=7, textColor=colors.HexColor("#AAAAAA"),
             fontName="Helvetica", alignment=TA_CENTER, spaceBefore=6)
    ))

    doc.build(story)
    buf.seek(0)
    return buf

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
    v3.0.0 · Gaso Comunicaciones<br>Procesamiento automático de inventario
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
            df_clean, df_consol, df_pending, logs, cols = run_pipeline(uploaded, saved_dec)

            st.session_state.df_clean      = df_clean
            st.session_state.df_consol     = df_consol
            st.session_state.df_pending    = df_pending
            st.session_state.logs       = logs
            st.session_state.cols       = cols
            st.session_state.processed  = True

            # Pre-build Excel
            uploaded.seek(0)
            excel_buf = build_excel_output(df_clean, df_consol, cols)
            st.session_state.excel_buf = excel_buf

        except Exception as e:
            st.error(f"❌ Error durante el procesamiento: {e}")
            import traceback; st.code(traceback.format_exc())

# ── Results ───────────────────────────────────────────────────────────────────
if st.session_state.processed:
    df_clean      = st.session_state.df_clean
    df_consol     = st.session_state.df_consol
    df_pending    = st.session_state.df_pending
    logs       = st.session_state.logs
    cols       = st.session_state.cols
    # Always ensure excel_buf exists — regenerate if the previous run failed
    if 'excel_buf' not in st.session_state or st.session_state.excel_buf is None:
        with st.spinner("Generando reporte Excel..."):
            st.session_state.excel_buf = build_excel_output(df_clean, df_consol, cols)
    excel_buf = st.session_state.excel_buf

    xdock_col  = cols["xdock"]
    pallet_col = cols["no_pallet"]

    # ── Processing log ────────────────────────────────────────────────────────
    with st.expander("📋 Log de Procesamiento", expanded=False):
        st.markdown('<div class="log-box">' + "<br>".join(logs) + "</div>",
                    unsafe_allow_html=True)


    # ── Filtro de Región ──────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">🗺️ Filtro de Región</div>', unsafe_allow_html=True)
    rf1, rf2, rf3 = st.columns([2, 2, 5])
    with rf1:
        region_filter = st.selectbox(
            "Analizar región",
            ["Todas", "REGIÓN LUIS", "REGIÓN JORGE"],
            key="region_filter",
            help="Filtra todos los KPIs, gráficas y tabla de ocupación por región."
        )
    with rf2:
        st.markdown("<br>", unsafe_allow_html=True)
        if region_filter != "Todas":
            xd_in_region = [xd for xd in CAPACIDADES if REGION_MAP.get(xd) == region_filter]
            st.caption(f"**{len(xd_in_region)} crossdocks** en {region_filter}: "
                       + ", ".join(CIUDAD_MAP.get(x,x) for x in xd_in_region))

    # Apply region filter to working dataframe for KPIs / charts / tables
    if region_filter == "Todas":
        df_view = df_clean.copy()
        xdocks_view = list(CAPACIDADES.keys())
    else:
        df_view = df_clean[df_clean["REGION"] == region_filter].copy()
        xdocks_view = [xd for xd in CAPACIDADES if REGION_MAP.get(xd) == region_filter]

    # ── KPIs ──────────────────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">📊 Indicadores Clave de Desempeño</div>',
                unsafe_allow_html=True)

    total_pal  = int(df_view[pallet_col].sum())
    total_m2   = round(df_view["M2"].sum(), 2)
    total_cap  = sum(CAPACIDADES[xd] for xd in xdocks_view)
    pct_global = round(total_m2 / total_cap * 100, 1) if total_cap > 0 else 0
    disponible = round(total_cap - total_m2, 2)
    n_pend     = len(df_pending)
    n_consol_kpi = len(df_consol[df_consol[xdock_col].isin(xdocks_view)] if len(df_consol)>0 else df_consol)

    kpi_cls = "kpi-red" if pct_global > 90 else "kpi-amber" if pct_global > 70 else "kpi-green"

    k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
    kpis_ui = [
        (k1, "Total Pallets",         f"{total_pal:,}",        "unidades",  "kpi-card"),
        (k2, "M² Ocupados",           f"{total_m2:,.1f}",      "m²",        "kpi-card"),
        (k3, "Capacidad Total",       f"{total_cap:,}",        "m²",        "kpi-card"),
        (k4, "% Ocupación Global",    f"{pct_global}%",        "",          f"kpi-card {kpi_cls}"),
        (k5, "M² Disponibles",        f"{disponible:,.1f}",    "m²",        "kpi-card"),
        (k6, "Sitios Consolidados",   f"{n_consol_kpi}",       "registros", "kpi-card" ),
        (k7, "Pendientes Revisión",   f"{n_pend}",             "registros",
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
    for xd in xdocks_view:
        cap    = CAPACIDADES[xd]
        m2_ocp = round(df_view[df_view[xdock_col] == xd]["M2"].sum(), 2)
        pal    = int(df_view[df_view[xdock_col] == xd][pallet_col].sum())
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
    # ── Sitios Consolidados ───────────────────────────────────────────────────
    n_consol = len(df_consol)
    st.markdown('<div class="sec-title">📌 Sitios Consolidados</div>', unsafe_allow_html=True)
    if n_consol > 0:
        xd_col = cols["xdock"]
        car_col = cols["carrier"]
        # Summary cards row
        consol_by_xd = df_consol.groupby(xd_col).size().reset_index(name="n")
        consol_by_car = df_consol.groupby(car_col).size().reset_index(name="n")
        ca, cb, cc = st.columns(3)
        with ca:
            st.markdown(f'''
            <div class="kpi-card" style="border-left-color:#8E44AD">
              <div class="kpi-label">Total Sitios Consolidados</div>
              <div class="kpi-value" style="color:#8E44AD">{n_consol}</div>
              <div class="kpi-unit">registros · sin impacto en capacidad</div>
            </div>''', unsafe_allow_html=True)
        with cb:
            top_xd = CIUDAD_MAP.get(consol_by_xd.sort_values("n", ascending=False).iloc[0][xd_col], "—") if len(consol_by_xd) > 0 else "—"
            top_n  = int(consol_by_xd["n"].max()) if len(consol_by_xd) > 0 else 0
            st.markdown(f'''
            <div class="kpi-card" style="border-left-color:#8E44AD">
              <div class="kpi-label">Crossdock con más consolidados</div>
              <div class="kpi-value" style="color:#8E44AD;font-size:1.3rem">{top_xd}</div>
              <div class="kpi-unit">{top_n} sitios</div>
            </div>''', unsafe_allow_html=True)
        with cc:
            carriers_c = ", ".join(sorted(df_consol[car_col].dropna().unique()))
            st.markdown(f'''
            <div class="kpi-card" style="border-left-color:#8E44AD">
              <div class="kpi-label">Carriers con consolidados</div>
              <div class="kpi-value" style="color:#8E44AD;font-size:1.1rem">{carriers_c}</div>
              <div class="kpi-unit">&nbsp;</div>
            </div>''', unsafe_allow_html=True)

        st.caption("Estos registros tienen **No. de Pallet = 0** — son materiales pequeños consolidados sobre otra tarima. "
                   "No se contabilizan en pallets ni en M² de capacidad.")

        # Summary table by XDOCK × Carrier
        grp_c = (df_consol.groupby([xd_col, car_col])
                 .agg(Sitios=(cols["id_sitio"], "count"),
                      ID_Sitios=(cols["id_sitio"], lambda x: ", ".join(x.dropna().unique()[:5])))
                 .reset_index())
        grp_c["Ciudad"] = grp_c[xd_col].map(CIUDAD_MAP)
        grp_c = grp_c[['Ciudad', xd_col, car_col, 'Sitios', 'ID_Sitios']]
        grp_c.columns = ['Ciudad', 'XDOCK', 'Carrier', 'Sitios Consolidados', 'ID Sitios (muestra)']
        st.dataframe(grp_c, use_container_width=True, height=min(200 + len(grp_c)*35, 380))
    else:
        st.info("No hay sitios consolidados en esta carga.")

    st.markdown('<div class="sec-title">📈 Dashboard Visual</div>', unsafe_allow_html=True)
    fig1, fig2, fig3, fig4, fig5, fig6, fig7 = make_charts(df_view, cols)

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

    # PDF generation
    st.markdown("**📄 Reporte PDF Ejecutivo**")
    pdf_r1, pdf_r2 = st.columns([3, 2])
    with pdf_r1:
        pdf_region = st.selectbox("Región para el PDF", ["Todas", "REGIÓN LUIS", "REGIÓN JORGE"],
                                   key="pdf_region_select")
    with pdf_r2:
        st.markdown("<br>", unsafe_allow_html=True)
        gen_pdf_btn = st.button("🖨️ Generar PDF Ejecutivo", use_container_width=True)

    if gen_pdf_btn:
        with st.spinner("Generando reporte PDF..."):
            try:
                pdf_buf = generate_pdf(df_clean, df_consol, cols, region_filter=pdf_region)
                st.session_state.pdf_buf = pdf_buf
                st.session_state.pdf_region = pdf_region
                st.success("✅ PDF generado correctamente")
            except Exception as e:
                st.error(f"Error generando PDF: {e}")
                import traceback; st.code(traceback.format_exc())

    if "pdf_buf" in st.session_state:
        region_label = st.session_state.get("pdf_region", "Todas").replace(" ","_").replace("Ó","O").replace("É","E")
        st.download_button(
            label="📥 Descargar PDF Ejecutivo",
            data=st.session_state.pdf_buf,
            file_name=f"GASO_REPORTE_EJECUTIVO_{region_label}_{fecha_str}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    st.markdown("---")
    st.markdown("**📊 Archivos Excel**")
    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button(
            label="📥 Reporte Completo Excel",
            data=excel_buf,
            file_name=f"GASO_REPORTE_{fecha_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d2:
        with st.spinner(""):
            orig_buf = build_original_format_excel(df_clean, df_consol, cols)
        st.download_button(
            label="📋 Base Limpia (formato original)",
            data=orig_buf,
            file_name=f"GASO_INOUT_LIMPIO_{fecha_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d3:
        csv_buf = df_clean.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="📄 Base Limpia CSV",
            data=csv_buf,
            file_name=f"GASO_INOUT_LIMPIO_{fecha_str}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # Footer
    st.markdown(f"""
    <div style="text-align:center;padding:1.2rem;margin-top:1.5rem;
                border-top:1px solid #DCE8F5;color:#AAA;font-size:.75rem">
      GASO COMUNICACIONES · IN-OUT Report Processor v3.0 ·
      Generado el {datetime.date.today().strftime('%d/%m/%Y')}
    </div>
    """, unsafe_allow_html=True)
