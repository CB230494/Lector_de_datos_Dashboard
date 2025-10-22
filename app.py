# -*- coding: utf-8 -*-
# ================================================================
# Lector de Matrices (Excel) → Resumen consolidado en Excel
# - Múltiples .xlsx/.xlsm
# - Detección por rótulos y contenido (layout-independiente)
# - GL/FP por conteo de filas; Avance por columnas "Avance" (trimestres)
# - Filtros anti falsos positivos (años, % como cantidades, etc.)
# - Descarga del consolidado a Excel
# ================================================================

import io, re, unicodedata
from typing import Dict, Optional, Tuple, List
import numpy as np
import pandas as pd
import streamlit as st

# ------------------------ Utilidades ----------------------------
def _norm(x: str) -> str:
    if x is None: return ""
    x = str(x)
    x = unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", x).strip().lower()

def _to_int(x) -> Optional[int]:
    """Solo acepta enteros de 1–3 dígitos (evita '2025' y descarta celdas con '%')."""
    if x is None: return None
    s = str(x).strip()
    if "%" in s:
        return None
    digits = re.sub(r"[^\d-]", "", s)
    if not re.fullmatch(r"-?\d{1,3}", digits):
        return None
    try:
        return int(digits)
    except:
        return None

def _to_pct(x) -> Optional[float]:
    """Devuelve porcentaje 0..100; si viene 0..1 lo escala; tolera '0', '0%'."""
    if x is None: return None
    s = str(x).replace(",", ".")
    m = re.search(r"-?\d+(\.\d+)?", s)
    if not m: return None
    v = float(m.group())
    if "%" in s or v > 1.0:
        return max(0.0, min(100.0, v))
    return max(0.0, min(100.0, v * 100.0))

def _read_df(file) -> pd.DataFrame:
    return pd.read_excel(file, engine="openpyxl", header=None, dtype=str)

def _find(df: pd.DataFrame, pattern: str) -> List[Tuple[int,int]]:
    rx = re.compile(pattern)
    out = []
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = df.iat[r,c]
            if val is None:
                continue
            if rx.search(_norm(val)):
                out.append((r,c))
    return out

def _neighbors(df: pd.DataFrame, r: int, c: int, up: int, down: int, left: int, right: int):
    r0 = max(0, r - up)
    r1 = min(df.shape[0]-1, r + down)
    c0 = max(0, c - left)
    c1 = min(df.shape[1]-1, c + right)
    for i in range(r0, r1+1):
        for j in range(c0, c1+1):
            yield i, j

def _pick_best_count(cands: List[int], max_allowed: int = 60) -> Optional[int]:
    cands = [x for x in cands if x is not None and 0 <= x <= max_allowed]
    if not cands:
        return None
    return max(cands)

# ------------------- Detectores por rótulos ---------------------
def detect_delegacion(df: pd.DataFrame) -> Optional[str]:
    rx = re.compile(r"^\s*d\d{1,3}\s*[-–]\s*.+\s*$", re.IGNORECASE)
    for r in range(min(15, df.shape[0])):
        for c in range(df.shape[1]):
            raw = df.iat[r,c]
            if raw and rx.match(str(raw)):
                return str(raw).strip()
    hits = _find(df, r"\bdelegaci[oó]n\b")
    for (r,c) in hits:
        if c+1 < df.shape[1]:
            v = df.iat[r, c+1]
            if v: return str(v).strip()
    return None

def detect_lineas_accion(df: pd.DataFrame, debug: bool=False) -> Optional[int]:
    hits = _find(df, r"\blineas?\s*de\s*accion\b")
    for (r,c) in hits:
        cands = []
        for (i,j) in _neighbors(df, r, c, up=0, down=6, left=2, right=4):
            cands.append(_to_int(df.iat[i,j]))
        val = _pick_best_count(cands, max_allowed=60)
        if debug and val is not None:
            st.caption(f"Líneas de Acción detectadas: {val}")
        if val is not None:
            return val
    la_hits = _find(df, r"\blinea\s+de\s+accion\s*#")
    if la_hits:
        return len(la_hits)
    return None

# -------- Detección robusta de GL/FP y Avance por “lo visible” -----
def detect_indicator_rows(df: pd.DataFrame) -> List[int]:
    rows = []
    for r in range(df.shape[0]):
        left_vals = [df.iat[r, c] for c in range(min(6, df.shape[1]))]
        left_norm = [_norm(v) for v in left_vals]
        if any(v == "gl" for v in left_norm) or any(v == "fp" for v in left_norm):
            rows.append(r)
    return rows

def detect_role_of_row(df: pd.DataFrame, r: int) -> Optional[str]:
    """Devuelve 'gl', 'fp' o None según el contenido a la izquierda."""
    left_vals = [df.iat[r, c] for c in range(min(6, df.shape[1]))]
    left_norm = [_norm(v) for v in left_vals]
    if any(v == "gl" for v in left_norm):
        return "gl"
    if any(v == "fp" for v in left_norm):
        return "fp"
    return None

def detect_avance_columns(df: pd.DataFrame) -> List[int]:
    cols = []
    for r in range(df.shape[0]):
        row = [df.iat[r, c] for c in range(df.shape[1])]
        for c, v in enumerate(row):
            if "avance" in _norm(v):
                cols.append(c)
        if len(cols) >= 2:
            return sorted(list(set(cols)))
        else:
            cols = []
    # Fallback razonable si no se detectan rótulos "Avance"
    fallback = [10, 15, 20, 25]
    return [c for c in fallback if c < df.shape[1]]

def gl_fp_counts(df: pd.DataFrame) -> Tuple[int, int]:
    rows = detect_indicator_rows(df)
    gl = fp = 0
    for r in rows:
        role = detect_role_of_row(df, r)
        if role == "gl":
            gl += 1
        elif role == "fp":
            fp += 1
    return gl, fp

def _row_status_from_avance(df: pd.DataFrame, r: int, avance_cols: List[int]) -> str:
    vals = [df.iat[r, c] for c in avance_cols]
    valsn = [_norm(v) for v in vals]
    if any("complet" in v for v in valsn):
        return "completos"
    if any("con actividades" in v for v in valsn):
        return "con_actividades"
    if any("sin actividades" in v for v in valsn):
        return "sin_actividades"
    return "sin_actividades"

def avance_counts(df: pd.DataFrame) -> Tuple[Dict[str,int], int]:
    """Conteo global (todas las filas de indicadores)."""
    rows = detect_indicator_rows(df)
    avance_cols = detect_avance_columns(df)
    counts = {"completos": 0, "con_actividades": 0, "sin_actividades": 0}
    for r in rows:
        counts[_row_status_from_avance(df, r, avance_cols)] += 1
    return counts, len(rows)

def avance_counts_by_role(df: pd.DataFrame) -> Tuple[Dict[str,int], int, Dict[str,int], int]:
    """Desglose por rol → (GL_counts, n_gl, FP_counts, n_fp)."""
    rows = detect_indicator_rows(df)
    avance_cols = detect_avance_columns(df)

    gl_counts = {"completos": 0, "con_actividades": 0, "sin_actividades": 0}
    fp_counts = {"completos": 0, "con_actividades": 0, "sin_actividades": 0}
    n_gl = n_fp = 0

    for r in rows:
        role = detect_role_of_row(df, r)
        stt = _row_status_from_avance(df, r, avance_cols)
        if role == "gl":
            gl_counts[stt] += 1
            n_gl += 1
        elif role == "fp":
            fp_counts[stt] += 1
            n_fp += 1

    return gl_counts, n_gl, fp_counts, n_fp

def detect_total_indicadores(df: pd.DataFrame) -> Optional[int]:
    hits = _find(df, r"\btotal\s+de\s+indicadores\b")
    for (r,c) in hits:
        cands = []
        for (i,j) in _neighbors(df, r, c, up=0, down=3, left=0, right=6):
            cands.append(_to_int(df.iat[i,j]))
        val = _pick_best_count([x for x in cands if x is not None], max_allowed=120)
        if val is not None:
            return val
    return None

# --------------------- Proceso de un archivo --------------------
def process_file(upload, debug: bool=False) -> Dict:
    df = _read_df(upload)

    deleg = detect_delegacion(df)
    lineas = detect_lineas_accion(df, debug=debug)

    gl, fp = gl_fp_counts(df)

    # Global
    avance_dict, total_ind = avance_counts(df)
    comp_n = avance_dict["completos"]
    con_n  = avance_dict["con_actividades"]
    sin_n  = avance_dict["sin_actividades"]

    def pct(n, d):
        return round((n / d) * 100.0, 1) if d and n is not None else None

    comp_p = pct(comp_n, total_ind)
    con_p  = pct(con_n,  total_ind)
    sin_p  = pct(sin_n,  total_ind)

    # --- NUEVO: desglose por GL y FP ---
    gl_counts, n_gl, fp_counts, n_fp = avance_counts_by_role(df)

    gl_comp_n = gl_counts["completos"]
    gl_con_n  = gl_counts["con_actividades"]
    gl_sin_n  = gl_counts["sin_actividades"]

    gl_comp_p = pct(gl_comp_n, n_gl)
    gl_con_p  = pct(gl_con_n,  n_gl)
    gl_sin_p  = pct(gl_sin_n,  n_gl)

    fp_comp_n = fp_counts["completos"]
    fp_con_n  = fp_counts["con_actividades"]
    fp_sin_n  = fp_counts["sin_actividades"]

    fp_comp_p = pct(fp_comp_n, n_fp)
    fp_con_p  = pct(fp_con_n,  n_fp)
    fp_sin_p  = pct(fp_sin_n,  n_fp)

    total_from_label = detect_total_indicadores(df)
    total_out = total_ind if total_ind else total_from_label

    # Ajuste si falta un lado (manteniendo tu lógica original)
    if (gl == 0 or fp == 0) and total_out and gl + fp != total_out:
        if gl == 0 and fp > 0:
            gl = max(0, total_out - fp)
        elif fp == 0 and gl > 0:
            fp = max(0, total_out - gl)

    out = {
        "archivo": upload.name,
        "delegacion": deleg,
        "lineas_accion": lineas,

        # Global (se conserva tal cual)
        "completos_n": comp_n,
        "completos_pct": comp_p,
        "conact_n": con_n,
        "conact_pct": con_p,
        "sinact_n": sin_n,
        "sinact_pct": sin_p,

        # Indicadores por rol (conteos)
        "indicadores_gl": gl if gl is not None else None,

        # --- NUEVAS 6 columnas GL (n y %) ---
        "gl_completos_n": gl_comp_n,
        "gl_completos_pct": gl_comp_p,
        "gl_conact_n": gl_con_n,
        "gl_conact_pct": gl_con_p,
        "gl_sinact_n": gl_sin_n,
        "gl_sinact_pct": gl_sin_p,

        "indicadores_fp": fp if fp is not None else None,

        # --- NUEVAS 6 columnas FP (n y %) ---
        "fp_completos_n": fp_comp_n,
        "fp_completos_pct": fp_comp_p,
        "fp_conact_n": fp_con_n,
        "fp_conact_pct": fp_con_p,
        "fp_sinact_n": fp_sin_n,
        "fp_sinact_pct": fp_sin_p,

        "indicadores_total": total_out if total_out is not None else (gl + fp if (gl or fp) else None),
    }

    if debug:
        st.caption(
            f"[DEBUG] Avance cols: {detect_avance_columns(df)} | rows GL/FP: {gl}+{fp}={gl+fp} | "
            f"n_gl={n_gl} n_fp={n_fp}"
        )
    return out

# --------------------------- UI --------------------------------
st.set_page_config(page_title="Lector de Matrices → Resumen Excel", layout="wide")
st.title("📊 Lector de Matrices (Excel) → Resumen consolidado")

with st.sidebar:
    st.header("Opciones")
    debug = st.toggle("Mostrar pistas de detección (debug)", value=False)

st.markdown("""
Sube tus matrices (.xlsx / .xlsm). La app detecta:
- **Delegación**, **Líneas de Acción**
- **Avance de Indicadores** (*Completos / Con actividades / Sin actividades*, con **n** y **%**, evaluado por fila de indicador y columnas **Avance**)
- **Indicadores** por **Gobierno Local** y **Fuerza Pública** (conteo de filas GL/FP) **y su desglose por estado (n y %)** ← *(nuevo)*
- **Total de Indicadores** (si existe)

y genera un **Excel consolidado** listo para descargar.
""")

uploads = st.file_uploader("Arrastra o selecciona tus matrices", type=["xlsx","xlsm"], accept_multiple_files=True)

rows, failed = [], []
if uploads:
    for f in uploads:
        try:
            rows.append(process_file(f, debug=debug))
        except Exception as e:
            failed.append((f.name, str(e)))

    if rows:
        df_out = pd.DataFrame(rows)

        # === Reordenación (se mantiene) + inserción de nuevas columnas ===
        rename = {
            "archivo":"Archivo",
            "delegacion":"Delegación",
            "lineas_accion":"Líneas de Acción",

            "indicadores_gl":"Indicadores Gobierno Local",

            # --- GL (nuevas 6) ---
            "gl_completos_n":"GL Completos (n)",
            "gl_completos_pct":"GL Completos (%)",
            "gl_conact_n":"GL Con actividades (n)",
            "gl_conact_pct":"GL Con actividades (%)",
            "gl_sinact_n":"GL Sin actividades (n)",
            "gl_sinact_pct":"GL Sin actividades (%)",

            "indicadores_fp":"Indicadores Fuerza Pública",

            # --- FP (nuevas 6) ---
            "fp_completos_n":"FP Completos (n)",
            "fp_completos_pct":"FP Completos (%)",
            "fp_conact_n":"FP Con actividades (n)",
            "fp_conact_pct":"FP Con actividades (%)",
            "fp_sinact_n":"FP Sin actividades (n)",
            "fp_sinact_pct":"FP Sin actividades (%)",

            "indicadores_total":"Total Indicadores",

            # Global (se conservan al final para referencia)
            "completos_n":"Completos (n)",
            "completos_pct":"Completos (%)",
            "conact_n":"Con actividades (n)",
            "conact_pct":"Con actividades (%)",
            "sinact_n":"Sin actividades (n)",
            "sinact_pct":"Sin actividades (%)",
        }

        order = [
            "archivo",
            "delegacion",
            "lineas_accion",

            "indicadores_gl",
            # --- aquí van las 6 de GL ---
            "gl_completos_n","gl_completos_pct",
            "gl_conact_n","gl_conact_pct",
            "gl_sinact_n","gl_sinact_pct",

            "indicadores_fp",
            # --- aquí van las 6 de FP ---
            "fp_completos_n","fp_completos_pct",
            "fp_conact_n","fp_conact_pct",
            "fp_sinact_n","fp_sinact_pct",

            "indicadores_total",

            # Global (al final, sin mover tu lógica)
            "completos_n","completos_pct",
            "conact_n","conact_pct",
            "sinact_n","sinact_pct",
        ]

        df_out = df_out[order].rename(columns=rename)

        # Formato %
        pct_cols = [
            "GL Completos (%)","GL Con actividades (%)","GL Sin actividades (%)",
            "FP Completos (%)","FP Con actividades (%)","FP Sin actividades (%)",
            "Completos (%)","Con actividades (%)","Sin actividades (%)",
        ]
        for col in pct_cols:
            if col in df_out.columns:
                df_out[col] = df_out[col].apply(lambda v: f"{v:.1f}%" if pd.notna(v) else None)

        st.subheader("Resumen previo")
        st.dataframe(df_out, use_container_width=True)

        # Descargar Excel
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df_out.to_excel(w, index=False, sheet_name="resumen")
        st.download_button(
            "⬇️ Descargar Excel consolidado",
            data=buf.getvalue(),
            file_name="resumen_matrices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    if failed:
        st.warning("Algunos archivos no se pudieron procesar automáticamente:")
        for name, err in failed:
            st.write(f"- {name}: {err}")
else:
    st.info("Sube tus matrices para ver el resumen.")

# ======================================================================
# =====================  PESTAÑA: 📊 Dashboard de Avance  ===============
# ======================================================================

import matplotlib.pyplot as plt

st.divider()
st.header("📊 Dashboard de Avance (extra)")

with st.expander("ℹ️ Instrucciones", expanded=True):
    st.markdown("""
    1) **Carga** aquí el **Excel consolidado** que genera esta misma app (hoja `resumen`).  
    2) Navega por las **pestañas**: *Por Delegación* y *Por Dirección Regional*.  
    3) Los colores y cuadros replican tu layout: **Rojo** (Sin actividades), **Amarillo** (Con actividades), **Verde** (Cumplida).
    """)

dash_file = st.file_uploader("Cargar Excel consolidado (resumen_matrices.xlsx)", type=["xlsx"], key="dash_excel")

# --------- helpers de parsing / estilos ----------
def _to_num_safe(x, pct=False):
    if pd.isna(x):
        return 0.0
    s = str(x).strip()
    if pct:
        s = s.replace("%", "").replace(",", ".")
        try:
            return float(s)
        except:
            return 0.0
    # enteros / floats en columnas (n)
    s = s.replace(",", ".")
    try:
        return float(s)
    except:
        # intenta extraer dígitos
        m = re.search(r"-?\d+(\.\d+)?", s)
        return float(m.group()) if m else 0.0

def _big_number(value, label, helptext=None):
    c = st.container()
    with c:
        st.markdown(f"""
        <div style="text-align:center;padding:8px 0;">
            <div style="font-size:54px;font-weight:800;line-height:1;margin:0;">{value}</div>
            <div style="font-size:14px;color:#555;margin-top:4px;">{label}</div>
        </div>
        """, unsafe_allow_html=True)
        if helptext:
            st.caption(helptext)

# Paleta
COLOR_ROJO   = "#ED1C24"
COLOR_AMARIL = "#F4C542"
COLOR_VERDE  = "#7AC943"
COLOR_AZUL_T = "#9BBBD9"
COLOR_AZUL_H = "#1F4E79"

def _bar_avance(pcts_tuple, title=""):
    # pcts_tuple: (sin%, con%, comp%)
    labels = ["Sin Actividades", "Con Actividades", "Cumplida"]
    values = list(pcts_tuple)
    colors = [COLOR_ROJO, COLOR_AMARIL, COLOR_VERDE]

    fig, ax = plt.subplots(figsize=(5.5, 3.5))
    ax.bar(labels, values, color=colors)
    ax.set_ylim(0, 100)
    ax.set_ylabel("%")
    ax.set_title(title)
    for i, v in enumerate(values):
        ax.text(i, v + 1, f"{v:.0f}%", ha="center", va="bottom", fontsize=10)
    st.pyplot(fig, use_container_width=True)

def _panel_tres(col, titulo, n_rojo, p_rojo, n_amar, p_amar, n_verde, p_verde, total):
    with col:
        # Header
        st.markdown(f"""
        <div style="background:{COLOR_AZUL_H};color:white;padding:8px 12px;border-radius:6px 6px 0 0;
                    font-weight:700;text-align:center;">{titulo}</div>
        """, unsafe_allow_html=True)
        # Body tipo tablero 3x2
        c1, c2, c3 = st.columns(3)
        c1.markdown(f"""
            <div style="background:{COLOR_ROJO};color:white;text-align:center;padding:8px;border:1px solid #ddd;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(n_rojo)}</div>
              <div style="font-size:13px;">Sin</div>
              <div style="font-size:16px;font-weight:700;">{p_rojo:.0f}%</div>
            </div>""", unsafe_allow_html=True)
        c2.markdown(f"""
            <div style="background:{COLOR_AMARIL};color:black;text-align:center;padding:8px;border:1px solid #ddd;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(n_amar)}</div>
              <div style="font-size:13px;">Con</div>
              <div style="font-size:16px;font-weight:700;">{p_amar:.0f}%</div>
            </div>""", unsafe_allow_html=True)
        c3.markdown(f"""
            <div style="background:{COLOR_VERDE};color:black;text-align:center;padding:8px;border:1px solid #ddd;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(n_verde)}</div>
              <div style="font-size:13px;">Cumplida</div>
              <div style="font-size:16px;font-weight:700;">{p_verde:.0f}%</div>
            </div>""", unsafe_allow_html=True)

        st.markdown(
            f"""<div style="text-align:center;border:1px solid #ddd;border-top:0;padding:8px 0;border-radius:0 0 6px 6px;">
            <div style="font-size:40px;font-weight:800;line-height:1;">{int(total)}</div>
            <div style="font-size:13px;color:#333;">Total</div></div>""",
            unsafe_allow_html=True
        )

def _resumen_avance(col, sin_n, sin_p, con_n, con_p, comp_n, comp_p, total_ind):
    with col:
        st.markdown(f"""
        <div style="background:{COLOR_AZUL_H};color:white;padding:8px 12px;border-radius:6px 6px 0 0;
                    font-weight:700;text-align:center;">Avance de Indicadores</div>
        """, unsafe_allow_html=True)

        # cuadricula
        grid = st.columns(3)
        grid[0].markdown(f"""
            <div style="background:{COLOR_ROJO};color:white;text-align:center;padding:8px;border:1px solid #ddd;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(sin_n)}</div>
              <div style="font-size:13px;">Sin Actividades</div>
              <div style="font-size:16px;font-weight:700;">{sin_p:.0f}%</div>
            </div>""", unsafe_allow_html=True)
        grid[1].markdown(f"""
            <div style="background:{COLOR_AMARIL};color:black;text-align:center;padding:8px;border:1px solid #ddd;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(con_n)}</div>
              <div style="font-size:13px;">Con Actividades</div>
              <div style="font-size:16px;font-weight:700;">{con_p:.0f}%</div>
            </div>""", unsafe_allow_html=True)
        grid[2].markdown(f"""
            <div style="background:{COLOR_VERDE};color:black;text-align:center;padding:8px;border:1px solid #ddd;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(comp_n)}</div>
              <div style="font-size:13px;">Cumplida</div>
              <div style="font-size:16px;font-weight:700;">{comp_p:.0f}%</div>
            </div>""", unsafe_allow_html=True)

        st.markdown(
            f"""<div style="text-align:center;border:1px solid #ddd;border-top:0;padding:8px 0;border-radius:0 0 6px 6px;">
            <div style="font-size:18px;font-weight:700;">Total de Indicadores: {int(total_ind)}</div></div>""",
            unsafe_allow_html=True
        )

def _ensure_numeric(df):
    # normaliza nombres esperados y garantiza numérico
    cols_n = [
        "GL Completos (n)","GL Con actividades (n)","GL Sin actividades (n)",
        "FP Completos (n)","FP Con actividades (n)","FP Sin actividades (n)",
        "Completos (n)","Con actividades (n)","Sin actividades (n)",
        "Indicadores Gobierno Local","Indicadores Fuerza Pública","Total Indicadores","Líneas de Acción"
    ]
    cols_p = [
        "GL Completos (%)","GL Con actividades (%)","GL Sin actividades (%)",
        "FP Completos (%)","FP Con actividades (%)","FP Sin actividades (%)",
        "Completos (%)","Con actividades (%)","Sin actividades (%)"
    ]
    for c in cols_n:
        if c in df.columns:
            df[c] = df[c].apply(_to_num_safe)
    for c in cols_p:
        if c in df.columns:
            df[c] = df[c].apply(lambda v: _to_num_safe(v, pct=True))
    return df

def _infer_dr(name: str) -> str:
    if not isinstance(name, str):
        return "Sin DR / No identificado"
    m = re.search(r"(DR-\s*\d+\S*)", name, flags=re.IGNORECASE)
    if m:
        return m.group(1).replace(" ", "")
    return "Sin DR / No identificado"

if dash_file:
    try:
        df_dash = pd.read_excel(dash_file, sheet_name="resumen")
    except Exception:
        # Si no trae nombre de hoja, intenta primera
        df_dash = pd.read_excel(dash_file)

    # Seguridad: columnas mínimas
    needed = [
        "Delegación","Líneas de Acción",
        "GL Completos (n)","GL Con actividades (n)","GL Sin actividades (n)",
        "GL Completos (%)","GL Con actividades (%)","GL Sin actividades (%)",
        "FP Completos (n)","FP Con actividades (n)","FP Sin actividades (n)",
        "FP Completos (%)","FP Con actividades (%)","FP Sin actividades (%)",
        "Completos (n)","Con actividades (n)","Sin actividades (n)",
        "Completos (%)","Con actividades (%)","Sin actividades (%)",
        "Indicadores Gobierno Local","Indicadores Fuerza Pública","Total Indicadores"
    ]
    # Aviso si falta algo (pero seguimos con lo que haya)
    faltantes = [c for c in needed if c not in df_dash.columns]
    if faltantes:
        st.warning("El Excel no trae algunas columnas esperadas (se mostrarán los paneles posibles): " + ", ".join(faltantes))

    df_dash = _ensure_numeric(df_dash.copy())
    df_dash["DR_inferida"] = df_dash["Delegación"].apply(_infer_dr)

    tabs = st.tabs(["🏢 Por Delegación", "🗺️ Por Dirección Regional"])

    # ======================= TAB 1: POR DELEGACIÓN =======================
    with tabs[0]:
        st.subheader("Avance por Delegación Policial")

        delegs = sorted(df_dash["Delegación"].dropna().astype(str).unique().tolist())
        if not delegs:
            st.info("No hay delegaciones en el archivo.")
        else:
            sel = st.selectbox("Delegación Policial", delegs, index=0, key="sel_deleg")

            # Si hay duplicados de la misma delegación, agregamos (sum)
            dsel = df_dash[df_dash["Delegación"] == sel]
            agg = dsel.select_dtypes(include=[np.number]).sum(numeric_only=True)
            # Para porcentajes globales, si vienen como % ya en columnas, usamos la primera fila válida;
            # si no, los recalculamos con los (n) y el total.
            def _pct_or_calc(n, tot, fallback):
                if not np.isnan(fallback) and fallback > 0:
                    return float(fallback)
                return (n / tot * 100.0) if tot > 0 else 0.0

            total_ind = agg.get("Total Indicadores", np.nan)
            if np.isnan(total_ind):
                total_ind = agg.get("Indicadores Gobierno Local", 0) + agg.get("Indicadores Fuerza Pública", 0)

            sin_n  = agg.get("Sin actividades (n)", 0)
            con_n  = agg.get("Con actividades (n)", 0)
            comp_n = agg.get("Completos (n)", 0)

            sin_p  = _pct_or_calc(sin_n, total_ind, dsel["Sin actividades (%)"].dropna().iloc[0] if "Sin actividades (%)" in dsel.columns and not dsel["Sin actividades (%)"].dropna().empty else np.nan)
            con_p  = _pct_or_calc(con_n, total_ind, dsel["Con actividades (%)"].dropna().iloc[0] if "Con actividades (%)" in dsel.columns and not dsel["Con actividades (%)"].dropna().empty else np.nan)
            comp_p = _pct_or_calc(comp_n, total_ind, dsel["Completos (%)"].dropna().iloc[0] if "Completos (%)" in dsel.columns and not dsel["Completos (%)"].dropna().empty else np.nan)

            # Título central y gráfico
            st.markdown(f"<h3 style='text-align:center;margin-top:0;'>{sel}</h3>", unsafe_allow_html=True)
            _bar_avance((sin_p, con_p, comp_p), title="Avance (%)")

            # Línea de acción en grande
            la_val = int(agg.get("Líneas de Acción", 0))
            _big_number(la_val, "Líneas de Acción")

            # Panel Izq: Avance de Indicadores (cuadros)
            col_left, col_right = st.columns([1.2, 1.2])

            _resumen_avance(
                col_left,
                sin_n, sin_p,
                con_n, con_p,
                comp_n, comp_p,
                total_ind
            )

            # Paneles Derecho: GL y FP
            gl_tot = agg.get("Indicadores Gobierno Local", 0)
            gl_sin_n  = agg.get("GL Sin actividades (n)", 0)
            gl_con_n  = agg.get("GL Con actividades (n)", 0)
            gl_comp_n = agg.get("GL Completos (n)", 0)
            gl_sin_p  = _pct_or_calc(gl_sin_n, gl_tot, dsel["GL Sin actividades (%)"].dropna().iloc[0] if "GL Sin actividades (%)" in dsel.columns and not dsel["GL Sin actividades (%)"].dropna().empty else np.nan)
            gl_con_p  = _pct_or_calc(gl_con_n, gl_tot, dsel["GL Con actividades (%)"].dropna().iloc[0] if "GL Con actividades (%)" in dsel.columns and not dsel["GL Con actividades (%)"].dropna().empty else np.nan)
            gl_comp_p = _pct_or_calc(gl_comp_n, gl_tot, dsel["GL Completos (%)"].dropna().iloc[0] if "GL Completos (%)" in dsel.columns and not dsel["GL Completos (%)"].dropna().empty else np.nan)

            fp_tot = agg.get("Indicadores Fuerza Pública", 0)
            fp_sin_n  = agg.get("FP Sin actividades (n)", 0)
            fp_con_n  = agg.get("FP Con actividades (n)", 0)
            fp_comp_n = agg.get("FP Completos (n)", 0)
            fp_sin_p  = _pct_or_calc(fp_sin_n, fp_tot, dsel["FP Sin actividades (%)"].dropna().iloc[0] if "FP Sin actividades (%)" in dsel.columns and not dsel["FP Sin actividades (%)"].dropna().empty else np.nan)
            fp_con_p  = _pct_or_calc(fp_con_n, fp_tot, dsel["FP Con actividades (%)"].dropna().iloc[0] if "FP Con actividades (%)" in dsel.columns and not dsel["FP Con actividades (%)"].dropna().empty else np.nan)
            fp_comp_p = _pct_or_calc(fp_comp_n, fp_tot, dsel["FP Completos (%)"].dropna().iloc[0] if "FP Completos (%)" in dsel.columns and not dsel["FP Completos (%)"].dropna().empty else np.nan)

            _panel_tres(
                col_right, "Gobierno Local",
                gl_sin_n, gl_sin_p, gl_con_n, gl_con_p, gl_comp_n, gl_comp_p, gl_tot
            )

            col_right2 = st.container()
            _panel_tres(
                st, "Fuerza Pública",
                fp_sin_n, fp_sin_p, fp_con_n, fp_con_p, fp_comp_n, fp_comp_p, fp_tot
            )

    # =================== TAB 2: POR DIRECCIÓN REGIONAL ===================
    with tabs[1]:
        st.subheader("Avance por Dirección Regional (DR)")
        drs = sorted(df_dash["DR_inferida"].unique().tolist())
        sel_dr = st.selectbox("Dirección Regional", drs, index=0, key="sel_dr")

        df_dr = df_dash[df_dash["DR_inferida"] == sel_dr]
        if df_dr.empty:
            st.info("No hay registros para esa DR.")
        else:
            # Agregados por DR
            agg = df_dr.select_dtypes(include=[np.number]).sum(numeric_only=True)

            total_ind = agg.get("Total Indicadores", np.nan)
            if np.isnan(total_ind):
                total_ind = agg.get("Indicadores Gobierno Local", 0) + agg.get("Indicadores Fuerza Pública", 0)

            sin_n  = agg.get("Sin actividades (n)", 0)
            con_n  = agg.get("Con actividades (n)", 0)
            comp_n = agg.get("Completos (n)", 0)

            def _pct(n): 
                return (n / total_ind * 100.0) if total_ind > 0 else 0.0

            sin_p, con_p, comp_p = _pct(sin_n), _pct(con_n), _pct(comp_n)

            st.markdown(f"<h3 style='text-align:center;margin-top:0;'>{sel_dr}</h3>", unsafe_allow_html=True)
            _bar_avance((sin_p, con_p, comp_p), title="Avance (%)")
            _big_number(int(agg.get("Líneas de Acción", 0)), "Líneas de Acción")

            col_left, col_right = st.columns([1.2, 1.2])
            _resumen_avance(col_left, sin_n, sin_p, con_n, con_p, comp_n, comp_p, total_ind)

            # GL y FP por DR
            gl_tot = agg.get("Indicadores Gobierno Local", 0)
            gl_sin_n  = agg.get("GL Sin actividades (n)", 0)
            gl_con_n  = agg.get("GL Con actividades (n)", 0)
            gl_comp_n = agg.get("GL Completos (n)", 0)
            gl_sin_p  = (gl_sin_n / gl_tot * 100.0) if gl_tot > 0 else 0.0
            gl_con_p  = (gl_con_n / gl_tot * 100.0) if gl_tot > 0 else 0.0
            gl_comp_p = (gl_comp_n / gl_tot * 100.0) if gl_tot > 0 else 0.0

            fp_tot = agg.get("Indicadores Fuerza Pública", 0)
            fp_sin_n  = agg.get("FP Sin actividades (n)", 0)
            fp_con_n  = agg.get("FP Con actividades (n)", 0)
            fp_comp_n = agg.get("FP Completos (n)", 0)
            fp_sin_p  = (fp_sin_n / fp_tot * 100.0) if fp_tot > 0 else 0.0
            fp_con_p  = (fp_con_n / fp_tot * 100.0) if fp_tot > 0 else 0.0
            fp_comp_p = (fp_comp_n / fp_tot * 100.0) if fp_tot > 0 else 0.0

            _panel_tres(
                col_right, "Gobierno Local",
                gl_sin_n, gl_sin_p, gl_con_n, gl_con_p, gl_comp_n, gl_comp_p, gl_tot
            )

            _panel_tres(
                st, "Fuerza Pública",
                fp_sin_n, fp_sin_p, fp_con_n, fp_con_p, fp_comp_n, fp_comp_p, fp_tot
            )
else:
    st.info("Carga el Excel consolidado para habilitar los dashboards.")


