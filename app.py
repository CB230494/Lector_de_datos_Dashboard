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
import matplotlib.pyplot as plt
from collections import Counter

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

    # Desglose por GL y FP
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

    # Ajuste si falta un lado
    if (gl == 0 or fp == 0) and total_out and gl + fp != total_out:
        if gl == 0 and fp > 0:
            gl = max(0, total_out - fp)
        elif fp == 0 and gl > 0:
            fp = max(0, total_out - gl)

    out = {
        "archivo": upload.name,
        "delegacion": deleg,
        "lineas_accion": lineas,

        # Global
        "completos_n": comp_n,
        "completos_pct": comp_p,
        "conact_n": con_n,
        "conact_pct": con_p,
        "sinact_n": sin_n,
        "sinact_pct": sin_p,

        # Indicadores por rol
        "indicadores_gl": gl if gl is not None else None,

        # GL (n y %)
        "gl_completos_n": gl_comp_n,
        "gl_completos_pct": gl_comp_p,
        "gl_conact_n": gl_con_n,
        "gl_conact_pct": gl_con_p,
        "gl_sinact_n": gl_sin_n,
        "gl_sinact_pct": gl_sin_p,

        "indicadores_fp": fp if fp is not None else None,

        # FP (n y %)
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
- **Avance de Indicadores** (*Completos / Con actividades / Sin actividades*, con **n** y **%**)
- **Indicadores** por **Gobierno Local** y **Fuerza Pública** (n y %)
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

        rename = {
            "archivo":"Archivo",
            "delegacion":"Delegación",
            "lineas_accion":"Líneas de Acción",

            "indicadores_gl":"Indicadores Gobierno Local",
            "gl_completos_n":"GL Completos (n)",
            "gl_completos_pct":"GL Completos (%)",
            "gl_conact_n":"GL Con actividades (n)",
            "gl_conact_pct":"GL Con actividades (%)",
            "gl_sinact_n":"GL Sin actividades (n)",
            "gl_sinact_pct":"GL Sin actividades (%)",

            "indicadores_fp":"Indicadores Fuerza Pública",
            "fp_completos_n":"FP Completos (n)",
            "fp_completos_pct":"FP Completos (%)",
            "fp_conact_n":"FP Con actividades (n)",
            "fp_conact_pct":"FP Con actividades (%)",
            "fp_sinact_n":"FP Sin actividades (n)",
            "fp_sinact_pct":"FP Sin actividades (%)",

            "indicadores_total":"Total Indicadores",

            "completos_n":"Completos (n)",
            "completos_pct":"Completos (%)",
            "conact_n":"Con actividades (n)",
            "conact_pct":"Con actividades (%)",
            "sinact_n":"Sin actividades (n)",
            "sinact_pct":"Sin actividades (%)",
        }

        order = [
            "archivo","delegacion","lineas_accion",
            "indicadores_gl",
            "gl_completos_n","gl_completos_pct","gl_conact_n","gl_conact_pct","gl_sinact_n","gl_sinact_pct",
            "indicadores_fp",
            "fp_completos_n","fp_completos_pct","fp_conact_n","fp_conact_pct","fp_sinact_n","fp_sinact_pct",
            "indicadores_total",
            "completos_n","completos_pct","conact_n","conact_pct","sinact_n","sinact_pct",
        ]

        df_out = df_out[order].rename(columns=rename)

        # Formato % (texto con %)
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

st.divider()
st.header("📊 Dashboard de Avance (extra)")

with st.expander("ℹ️ Instrucciones", expanded=True):
    st.markdown("""
    1) **Carga** aquí el **Excel consolidado** (hoja `resumen`).  
    2) Pestañas: **Por Delegación**, **Por Dirección Regional**, y **Gobierno Local (por Provincia)**.  
    3) En todas se muestra **Líneas de Acción total** y, si el archivo trae, el **desglose**: Gobierno Local / Fuerza Pública / Mixtas.  
    4) La primera sección de la app (consolidado desde múltiples matrices) sigue igual; este dashboard usa ese mismo **único Excel**.
    """)

dash_file = st.file_uploader("Cargar Excel consolidado (resumen_matrices.xlsx)", type=["xlsx"], key="dash_excel")

# -------------------- helpers de parsing / estilos ---------------------
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
    s = s.replace(",", ".")
    try:
        return float(s)
    except:
        m = re.search(r"-?\d+(\.\d+)?", s)
        return float(m.group()) if m else 0.0

def _big_number(value, label, helptext=None, big_px=80):
    c = st.container()
    with c:
        st.markdown(f"""
        <div style="text-align:center;padding:10px;background:#ffffff;border:1px solid #e3e3e3;border-radius:8px;">
            <div style="font-size:14px;color:#666;margin-bottom:6px;">{label}</div>
            <div style="font-size:{big_px}px;font-weight:900;line-height:1;color:#111;">{value}</div>
        </div>
        """, unsafe_allow_html=True)
        if helptext:
            st.caption(helptext)

# Paleta
COLOR_ROJO   = "#ED1C24"
COLOR_AMARIL = "#F4C542"
COLOR_VERDE  = "#7AC943"
COLOR_AZUL_H = "#1F4E79"

# === Gráfico con fondo BLANCO ===
def _bar_avance(pcts_tuple, title=""):
    labels = ["Sin actividades", "Con actividades", "Cumplida"]
    values = list(pcts_tuple)
    colors = [COLOR_ROJO, COLOR_AMARIL, COLOR_VERDE]
    fig, ax = plt.subplots(figsize=(5.5, 3.5))
    fig.patch.set_facecolor("#ffffff")
    ax.set_facecolor("#ffffff")
    ax.bar(labels, values, color=colors)
    ax.set_ylim(0, 100)
    ax.set_ylabel("%", color="#111")
    ax.set_title(title, color="#111")
    ax.tick_params(axis="x", colors="#111")
    ax.tick_params(axis="y", colors="#111")
    for spine in ax.spines.values():
        spine.set_color("#999")
    for i, v in enumerate(values):
        ax.text(i, v + 1, f"{v:.0f}%", ha="center", va="bottom", fontsize=10, color="#111")
    st.pyplot(fig, use_container_width=True)

def _panel_tres(col, titulo, n_rojo, p_rojo, n_amar, p_amar, n_verde, p_verde, total):
    with col:
        st.markdown(f"""
        <div style="background:{COLOR_AZUL_H};color:white;padding:10px 12px;border-radius:8px 8px 0 0;
                    font-weight:700;text-align:center;border:1px solid #e3e3e3;border-bottom:0;">{titulo}</div>
        """, unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        c1.markdown(f"""
            <div style="background:{COLOR_ROJO};color:white;text-align:center;padding:10px;border:1px solid #e3e3e3;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(n_rojo)}</div>
              <div style="font-size:13px;">Sin actividades</div>
              <div style="font-size:16px;font-weight:700;">{p_rojo:.0f}%</div>
            </div>""", unsafe_allow_html=True)
        c2.markdown(f"""
            <div style="background:{COLOR_AMARIL};color:#111;text-align:center;padding:10px;border:1px solid #e3e3e3;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(n_amar)}</div>
              <div style="font-size:13px;">Con actividades</div>
              <div style="font-size:16px;font-weight:700;">{p_amar:.0f}%</div>
            </div>""", unsafe_allow_html=True)
        c3.markdown(f"""
            <div style="background:{COLOR_VERDE};color:#111;text-align:center;padding:10px;border:1px solid #e3e3e3;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(n_verde)}</div>
              <div style="font-size:13px;">Cumplida</div>
              <div style="font-size:16px;font-weight:700;">{p_verde:.0f}%</div>
            </div>""", unsafe_allow_html=True)

        st.markdown(
            f"""<div style="text-align:center;border:1px solid #e3e3e3;border-top:0;padding:10px;border-radius:0 0 8px 8px;background:#ffffff;color:#111;">
            <div style="font-size:13px;color:#666;margin-bottom:4px;">Total de indicadores</div>
            <div style="font-size:44px;font-weight:900;line-height:1;">{int(total)}</div></div>""",
            unsafe_allow_html=True
        )

def _resumen_avance(col, sin_n, sin_p, con_n, con_p, comp_n, comp_p, total_ind):
    with col:
        st.markdown(f"""
        <div style="background:{COLOR_AZUL_H};color:white;padding:10px 12px;border-radius:8px 8px 0 0;
                    font-weight:700;text-align:center;border:1px solid #e3e3e3;border-bottom:0;">Avance de Indicadores</div>
        """, unsafe_allow_html=True)

        grid = st.columns(3)
        grid[0].markdown(f"""
            <div style="background:{COLOR_ROJO};color:white;text-align:center;padding:10px;border:1px solid #e3e3e3;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(sin_n)}</div>
              <div style="font-size:13px;">Sin actividades</div>
              <div style="font-size:16px;font-weight:700;">{sin_p:.0f}%</div>
            </div>""", unsafe_allow_html=True)
        grid[1].markdown(f"""
            <div style="background:{COLOR_AMARIL};color:#111;text-align:center;padding:10px;border:1px solid #e3e3e3;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(con_n)}</div>
              <div style="font-size:13px;">Con actividades</div>
              <div style="font-size:16px;font-weight:700;">{con_p:.0f}%</div>
            </div>""", unsafe_allow_html=True)
        grid[2].markdown(f"""
            <div style="background:{COLOR_VERDE};color:#111;text-align:center;padding:10px;border:1px solid #e3e3e3;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(comp_n)}</div>
              <div style="font-size:13px;">Cumplida</div>
              <div style="font-size:16px;font-weight:700;">{comp_p:.0f}%</div>
            </div>""", unsafe_allow_html=True)

        st.markdown(
            f"""<div style="text-align:center;border:1px solid #e3e3e3;border-top:0;padding:14px;border-radius:0 0 8px 8px;background:#ffffff;color:#111;">
            <div style="font-size:13px;color:#666;margin-bottom:6px;">Total de indicadores (Gobierno Local + Fuerza Pública)</div>
            <div style="font-size:60px;font-weight:900;line-height:1;">{int(total_ind)}</div></div>""",
            unsafe_allow_html=True
        )

def _ensure_numeric(df):
    cols_n = [
        "GL Completos (n)","GL Con actividades (n)","GL Sin actividades (n)",
        "FP Completos (n)","FP Con actividades (n)","FP Sin actividades (n)",
        "Completos (n)","Con actividades (n)","Sin actividades (n)",
        "Indicadores Gobierno Local","Indicadores Fuerza Pública","Total Indicadores",
        "Líneas de Acción","Líneas de Acción Gobierno Local","Líneas de Acción Fuerza Pública","Líneas de Acción Mixtas"
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

# ------------------ DR / Provincia: detección robusta -------------------
def _norm_str(s: str) -> str:
    if s is None: return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    s = re.sub(r"[\s_\-]+", " ", s).strip().lower()
    return s

def _infer_dr_from_delegacion(name: str) -> str:
    if not isinstance(name, str):
        return "Sin DR / No identificado"
    m = re.search(r"(R\s*\d+)", name, flags=re.IGNORECASE)
    if m:
        return m.group(1).upper().replace(" ", "")
    m = re.search(r"(DR-\s*\d+\S*)", name, flags=re.IGNORECASE)
    if m:
        return m.group(1).replace(" ", "")
    return "Sin DR / No identificado"

def _pick_dr_column(df: pd.DataFrame) -> Optional[str]:
    norm_map = {_norm_str(c): c for c in df.columns}
    candidates = [
        "direccionregional","direccion regional","dirregional",
        "dr","region","regional","direccion","direccionreg","regiones"
    ]
    for cand in candidates:
        if cand in norm_map:
            return norm_map[cand]
    for k, real in norm_map.items():
        if ("direccion" in k and "regional" in k) or k == "dr":
            return real
    return None

def _dr_sort_key(s: str):
    if not isinstance(s, str):
        return (999, "")
    m = re.search(r"r\s*([0-9]+)", s, flags=re.IGNORECASE)
    num = int(m.group(1)) if m else 999
    return (num, s)

def _pick_prov_column(df: pd.DataFrame) -> Optional[str]:
    norm_map = {_norm_str(c): c for c in df.columns}
    candidates = ["provincia", "province", "prov", "provincial", "provincia nombre", "nom provincia"]
    for cand in candidates:
        if cand in norm_map:
            return norm_map[cand]
    for k, real in norm_map.items():
        if "provinc" in k:
            return real
    return None

# --------- helper: total y desglose de Líneas de Acción ---------------
def _lineas_tot_y_desglose(agg: pd.Series):
    has_gl = "Líneas de Acción Gobierno Local" in agg.index
    has_fp = "Líneas de Acción Fuerza Pública" in agg.index
    has_mx = "Líneas de Acción Mixtas" in agg.index
    if has_gl or has_fp or has_mx:
        gl = int(agg.get("Líneas de Acción Gobierno Local", 0) or 0)
        fp = int(agg.get("Líneas de Acción Fuerza Pública", 0) or 0)
        mx = int(agg.get("Líneas de Acción Mixtas", 0) or 0)
        total = gl + fp + mx
        return total, gl, fp, mx, True
    total = int(agg.get("Líneas de Acción", 0) or 0)
    return total, None, None, None, False

def _render_lineas_block(agg: pd.Series):
    total, gl, fp, mx, has_breakdown = _lineas_tot_y_desglose(agg)
    _big_number(total, "Líneas de Acción", big_px=72)
    if has_breakdown:
        c1, c2, c3 = st.columns(3)
        c1.markdown(f"""
            <div style="background:#ffffff;border:1px solid #e3e3e3;border-radius:8px;padding:10px;text-align:center;">
              <div style="font-size:13px;color:#666;margin-bottom:4px;">Gobierno Local</div>
              <div style="font-size:32px;font-weight:800;color:#111;">{gl}</div>
            </div>""", unsafe_allow_html=True)
        c2.markdown(f"""
            <div style="background:#ffffff;border:1px solid #e3e3e3;border-radius:8px;padding:10px;text-align:center;">
              <div style="font-size:13px;color:#666;margin-bottom:4px;">Fuerza Pública</div>
              <div style="font-size:32px;font-weight:800;color:#111;">{fp}</div>
            </div>""", unsafe_allow_html=True)
        c3.markdown(f"""
            <div style="background:#ffffff;border:1px solid #e3e3e3;border-radius:8px;padding:10px;text-align:center;">
              <div style="font-size:13px;color:#666;margin-bottom:4px;">Mixtas</div>
              <div style="font-size:32px;font-weight:800;color:#111;">{mx}</div>
            </div>""", unsafe_allow_html=True)

# --------- toggle: suma TODAS las opciones del filtro ---------------
def _scope_total_o_seleccion(df_selected: pd.DataFrame, df_all_options: pd.DataFrame, key: str, etiqueta_plural: str):
    use_total = st.toggle("Habilitar mostrar total de datos", value=False, key=key)
    scope = df_all_options if use_total else df_selected
    st.markdown(
        f"""
        <div style="margin:8px 0 0 0; text-align:center;">
          <span style="display:inline-block;padding:8px 14px;border-radius:999px;border:1px solid #e3e3e3;background:#ffffff;color:#111;">
            <span style="font-size:13px;color:#666;margin-right:8px;">Total de datos del filtro</span>
            <span style="font-size:24px;font-weight:900;line-height:1;">{len(scope)}</span>
          </span>
        </div>
        """,
        unsafe_allow_html=True
    )
    if use_total:
        st.caption(f"Mostrando la **totalidad** de {etiqueta_plural}.")
    return scope, use_total

# ==================== NUEVO: Filtros + resumen por Categorías/Problemática/Responsable ====================
def _pick_named_column(df: pd.DataFrame, names_like: List[str]) -> Optional[str]:
    if df is None or df.empty: return None
    norm_map = {_norm_str(c): c for c in df.columns}
    for alias in names_like:
        a = _norm_str(alias)
        if a in norm_map:
            return norm_map[a]
    for k, real in norm_map.items():
        if any(_norm_str(x) in k for x in names_like):
            return real
    return None

def _extract_tokens(cell) -> List[str]:
    if pd.isna(cell):
        return []
    s = str(cell)
    parts = re.split(r"[,\;/\|\n]+", s)
    out = []
    for p in parts:
        p2 = re.sub(r"\s+", " ", p).strip()
        if p2:
            out.append(p2)
    return out

def _role_flags(row: pd.Series) -> Tuple[bool, bool, bool]:
    la_gl = row.get("Líneas de Acción Gobierno Local", np.nan)
    la_fp = row.get("Líneas de Acción Fuerza Pública", np.nan)
    la_mx = row.get("Líneas de Acción Mixtas", np.nan)
    ind_gl = row.get("Indicadores Gobierno Local", np.nan)
    ind_fp = row.get("Indicadores Fuerza Pública", np.nan)

    has_la = not (pd.isna(la_gl) and pd.isna(la_fp) and pd.isna(la_mx))
    if has_la:
        gl = _to_num_safe(la_gl) > 0 if not pd.isna(la_gl) else False
        fp = _to_num_safe(la_fp) > 0 if not pd.isna(la_fp) else False
        mx = _to_num_safe(la_mx) > 0 if not pd.isna(la_mx) else False
        if not mx and (gl and fp):
            mx = True
    else:
        gl = _to_num_safe(ind_gl) > 0 if not pd.isna(ind_gl) else False
        fp = _to_num_safe(ind_fp) > 0 if not pd.isna(ind_fp) else False
        mx = gl and fp
    return gl, fp, mx

# --- NUEVO: sumar líneas de acción (total y por rol) sobre un DF filtrado ---
def _sum_lineas(df_any: pd.DataFrame) -> Tuple[int, int, int, int]:
    """Devuelve (total, gl, fp, mx) sumando columnas de líneas de acción si existen; si no, infiere."""
    if df_any is None or df_any.empty:
        return 0, 0, 0, 0
    num = df_any.select_dtypes(include=[np.number])
    if {"Líneas de Acción Gobierno Local","Líneas de Acción Fuerza Pública","Líneas de Acción Mixtas"}.issubset(set(num.columns)):
        gl = int(num["Líneas de Acción Gobierno Local"].sum())
        fp = int(num["Líneas de Acción Fuerza Pública"].sum())
        mx = int(num["Líneas de Acción Mixtas"].sum())
        return gl + fp + mx, gl, fp, mx
    if "Líneas de Acción" in num.columns:
        tot = int(num["Líneas de Acción"].sum())
        return tot, None or 0, None or 0, None or 0
    # último recurso: contar filas (no ideal, pero evita romper)
    return len(df_any), 0, 0, 0

def _extra_filters_ui(scope_df: pd.DataFrame, key_prefix: str = "") -> Tuple[pd.DataFrame, Dict[str, str]]:
    if scope_df is None or scope_df.empty:
        return scope_df, {}

    col_categorias = _pick_named_column(scope_df, ["categorias", "categoria", "categoría", "categorías"])
    col_problema   = _pick_named_column(scope_df, ["problematica", "problemática", "problema", "problemas"])
    col_resp       = _pick_named_column(scope_df, ["responsable", "responsables"])

    found = {k:v for k,v in {
        "Categorias": col_categorias,
        "Problematica": col_problema,
        "Responsable": col_resp
    }.items() if v is not None}

    st.markdown("### 🔎 Filtro adicional (Categorías / Problemática / Responsable)")
    if not found:
        st.caption("No se encontraron columnas **Categorias**, **Problematica** o **Responsable** en el Excel. (Se muestran solo los paneles superiores).")
        return scope_df, {}

    c1, c2, c3 = st.columns(3)
    df_f = scope_df.copy()

    if col_categorias:
        opts = sorted(pd.unique(df_f[col_categorias].dropna().astype(str)))
        sel = c1.multiselect("Categorias", options=opts, default=opts, key=f"{key_prefix}cat")
        if sel and len(sel) < len(opts):
            df_f = df_f[df_f[col_categorias].astype(str).isin(sel)]

    if col_problema:
        opts2 = sorted(pd.unique(df_f[col_problema].dropna().astype(str)))
        sel2 = c2.multiselect("Problemática", options=opts2, default=opts2, key=f"{key_prefix}prob")
        if sel2 and len(sel2) < len(opts2):
            df_f = df_f[df_f[col_problema].astype(str).isin(sel2)]

    if col_resp:
        opts3 = sorted(pd.unique(df_f[col_resp].dropna().astype(str)))
        sel3 = c3.multiselect("Responsable", options=opts3, default=opts3, key=f"{key_prefix}resp")
        if sel3 and len(sel3) < len(opts3):
            df_f = df_f[df_f[col_resp].astype(str).isin(sel3)]

    return df_f, found

def _render_informativo(df_filt: pd.DataFrame, found_cols: Dict[str, str], key_prefix: str = ""):
    st.markdown("---")
    st.markdown("### 🧾 Resumen informativo (nuevo)")

    if df_filt is None or df_filt.empty:
        st.info("No hay registros que coincidan con los filtros seleccionados.")
        return

    col_categorias = found_cols.get("Categorias")
    col_problema   = found_cols.get("Problematica")

    if not col_categorias and not col_problema:
        st.caption("No hay columnas de Categorías o Problemática para resumir.")
        return

    # Contadores por rol
    cats_gl = Counter(); cats_fp = Counter(); cats_mx = Counter()
    probs_gl = Counter(); probs_fp = Counter(); probs_mx = Counter()

    delito_gl = Counter(); delito_fp = Counter(); delito_mx = Counter()
    riesgo_gl = Counter(); riesgo_fp = Counter(); riesgo_mx = Counter()

    # Para listado “riesgos sociales y quién los atiende”
    # mapeo problemática -> set(roles)
    riesgo_roles = {}

    for _, row in df_filt.iterrows():
        gl, fp, mx = _role_flags(row)
        cats = _extract_tokens(row[col_categorias]) if col_categorias else []
        probs = _extract_tokens(row[col_problema]) if col_problema else []
        cats_norm = [_norm_str(c) for c in cats]

        if gl:
            for c in cats:  cats_gl[c] += 1
            for p in probs: probs_gl[p] += 1
        if fp:
            for c in cats:  cats_fp[c] += 1
            for p in probs: probs_fp[p] += 1
        if mx:
            for c in cats:  cats_mx[c] += 1
            for p in probs: probs_mx[p] += 1

        # Marcadores especiales por categoria
        is_delito = any(k in cats_norm for k in ["delito","delitos"])
        is_riesgo = any(k in cats_norm for k in ["riesgo social","riesgos sociales","riesgo","riesgos"])

        if is_delito:
            if gl: 
                for p in probs: delito_gl[p] += 1
            if fp: 
                for p in probs: delito_fp[p] += 1
            if mx: 
                for p in probs: delito_mx[p] += 1

        if is_riesgo:
            if gl: 
                for p in probs: riesgo_gl[p] += 1
            if fp: 
                for p in probs: riesgo_fp[p] += 1
            if mx: 
                for p in probs: riesgo_mx[p] += 1
            # Registrar roles para el listado por problemática
            for p in probs:
                roles = riesgo_roles.get(p, set())
                if gl: roles.add("Gobierno Local")
                if fp: roles.add("Fuerza Pública")
                if mx: roles.add("Mixta")
                riesgo_roles[p] = roles

    def _render_counter(title, cnt: Counter):
        items = cnt.most_common()
        if not items:
            st.caption(f"• Sin datos para **{title}**.")
            return
        lis = "".join([f"<li><b>{k}</b>: {v}</li>" for k,v in items])
        st.markdown(f"""
        <div style="background:#fff;border:1px solid #e3e3e3;border-radius:8px;padding:10px;">
          <div style="font-weight:700;margin-bottom:6px;color:#111;">{title}</div>
          <ul style="margin:0 0 0 18px;color:#111;">{lis}</ul>
        </div>
        """, unsafe_allow_html=True)

    # Bloques: Categorías por rol
    st.markdown("#### Categorías atendidas por rol")
    c1, c2, c3 = st.columns(3)
    with c1: _render_counter("Gobierno Local", cats_gl)
    with c2: _render_counter("Fuerza Pública", cats_fp)
    with c3: _render_counter("Mixta", cats_mx)

    # Bloques: Problemáticas por rol
    st.markdown("#### Problemáticas atendidas por rol")
    p1, p2, p3 = st.columns(3)
    with p1: _render_counter("Gobierno Local", probs_gl)
    with p2: _render_counter("Fuerza Pública", probs_fp)
    with p3: _render_counter("Mixta", probs_mx)

    # ===================== NUEVO RESUMEN TEXTUAL QUE PEDISTE =====================
    # 1) "Se atienden X líneas de acción"
    total_lineas, gl_lineas, fp_lineas, mx_lineas = _sum_lineas(df_filt)
    st.markdown("#### 🧮 Resumen de líneas de acción")
    st.markdown(
        f"<div style='background:#fff;border:1px solid #e3e3e3;border-radius:8px;padding:12px;color:#111;'>"
        f"<div><b>Se atienden {int(total_lineas)} líneas de acción</b></div>"
        + (
            f"<div style='margin-top:6px;'>Desglose: "
            f"Gobierno Local: <b>{int(gl_lineas)}</b> &nbsp;•&nbsp; "
            f"Fuerza Pública: <b>{int(fp_lineas)}</b> &nbsp;•&nbsp; "
            f"Mixtas: <b>{int(mx_lineas)}</b></div>"
            if (gl_lineas or fp_lineas or mx_lineas) else ""
        )
        + "</div>",
        unsafe_allow_html=True
    )

    # 2) "De esas, Y corresponden a Riesgos sociales (GL:a, FP:b, Mixtas:c)"
    #    Sumamos líneas solo en filas cuya Categoría sea riesgo social
    if col_categorias:
        mask_riesgo = df_filt[col_categorias].astype(str).str.lower().str.contains(
            r"(riesgo social|riesgos sociales|riesgo\b|riesgos\b)", regex=True, na=False
        )
        df_riesgo = df_filt[mask_riesgo].copy()
        r_tot, r_gl, r_fp, r_mx = _sum_lineas(df_riesgo)
        st.markdown(
            f"<div style='background:#fff;border:1px solid #e3e3e3;border-radius:8px;padding:12px;margin-top:8px;color:#111;'>"
            f"De esas, <b>{int(r_tot)}</b> corresponden a <b>Riesgos sociales</b>"
            + (
                f" (GL: <b>{int(r_gl)}</b>, FP: <b>{int(r_fp)}</b>, Mixtas: <b>{int(r_mx)}</b>)"
                if (r_gl or r_fp or r_mx) else ""
            )
            + "</div>",
            unsafe_allow_html=True
        )
    else:
        df_riesgo = pd.DataFrame()

    # 3) “Riesgos sociales y quién los atiende” (lista)
    st.markdown("#### 📋 Riesgos sociales y quién los atiende")
    if not riesgo_roles:
        st.caption("No se identificaron problemáticas bajo la categoría **Riesgo social** en el alcance seleccionado.")
    else:
        items = sorted(riesgo_roles.items(), key=lambda x: x[0].lower())
        lis = []
        for problema, roles in items:
            roles_txt = ", ".join(sorted(roles))
            lis.append(f"<li><b>{problema}</b> — {roles_txt}</li>")
        st.markdown(
            "<div style='background:#fff;border:1px solid #e3e3e3;border-radius:8px;padding:12px;'>"
            "<ul style='margin:0 0 0 18px;color:#111;'>"
            + "".join(lis) +
            "</ul></div>",
            unsafe_allow_html=True
        )

    # ----------------- Secciones especiales (ya las tenías) -----------------
    st.markdown("#### Cuando la **Categoría** es “Delito” o “Riesgo Social”")
    st.caption("Se listan las **problemáticas** asociadas y sus conteos por rol.")

    st.markdown("**Delito**")
    d1, d2, d3 = st.columns(3)
    with d1: _render_counter("Gobierno Local", delito_gl)
    with d2: _render_counter("Fuerza Pública", delito_fp)
    with d3: _render_counter("Mixta", delito_mx)

    st.markdown("**Riesgo Social**")
    r1, r2, r3 = st.columns(3)
    with r1: _render_counter("Gobierno Local", riesgo_gl)
    with r2: _render_counter("Fuerza Pública", riesgo_fp)
    with r3: _render_counter("Mixta", riesgo_mx)

# ============================= MAIN DASHBOARD =============================
if dash_file:
    try:
        df_dash = pd.read_excel(dash_file, sheet_name="resumen")
    except Exception:
        df_dash = pd.read_excel(dash_file)

    df_dash = _ensure_numeric(df_dash.copy())

    # DR inferida
    dr_col = _pick_dr_column(df_dash)
    if dr_col:
        tmp = (
            df_dash[dr_col].astype(str).apply(_norm_str)
            .str.replace(r"\s+", " ", regex=True).str.strip()
        )
        tmp = tmp.str.replace(
            r"(^r\s*\d+)\s*", lambda m: m.group(1).upper().replace(" ", "") + " ", regex=True
        )
        tmp = tmp.replace({"": "Sin DR / No identificado", "nan": "Sin DR / No identificado", "none": "Sin DR / No identificado"})
        df_dash["DR_inferida"] = tmp
    else:
        df_dash["DR_inferida"] = df_dash.get("Delegación", "").apply(_infer_dr_from_delegacion)

    tabs = st.tabs(["🏢 Por Delegación", "🗺️ Por Dirección Regional", "🏛️ Gobierno Local (por Provincia)"])

    # ======================= TAB 1: POR DELEGACIÓN =======================
    with tabs[0]:
        st.subheader("Avance por Delegación Policial")

        if "Delegación" not in df_dash.columns:
            st.info("El Excel no contiene la columna 'Delegación'.")
        else:
            delegs = sorted(df_dash["Delegación"].dropna().astype(str).unique().tolist())
            sel = st.selectbox("Delegación Policial", delegs, index=0, key="sel_deleg")

            dsel = df_dash[df_dash["Delegación"] == sel]

            scope_df, using_total = _scope_total_o_seleccion(
                df_selected=dsel,
                df_all_options=df_dash,
                key="toggle_total_deleg",
                etiqueta_plural="delegaciones"
            )

            agg = scope_df.select_dtypes(include=[np.number]).sum(numeric_only=True)

            total_ind = agg.get("Total Indicadores", np.nan)
            if np.isnan(total_ind):
                total_ind = agg.get("Indicadores Gobierno Local", 0) + agg.get("Indicadores Fuerza Pública", 0)

            sin_n  = agg.get("Sin actividades (n)", 0)
            con_n  = agg.get("Con actividades (n)", 0)
            comp_n = agg.get("Completos (n)", 0)

            def _pct(n, d): 
                return (n / d * 100.0) if d > 0 else 0.0

            sin_p, con_p, comp_p = _pct(sin_n, total_ind), _pct(con_n, total_ind), _pct(comp_n, total_ind)

            titulo_h3 = "Total (todas las delegaciones)" if using_total else sel
            st.markdown(f"<h3 style='text-align:center;margin-top:0;color:#111;'>{titulo_h3}</h3>", unsafe_allow_html=True)

            _render_lineas_block(agg)
            _bar_avance((sin_p, con_p, comp_p), title="Total de indicadores (Gobierno Local + Fuerza Pública)")

            top_gl, top_fp = st.columns(2)
            gl_tot = agg.get("Indicadores Gobierno Local", 0)
            gl_sin_n  = agg.get("GL Sin actividades (n)", 0); gl_con_n  = agg.get("GL Con actividades (n)", 0); gl_comp_n = agg.get("GL Completos (n)", 0)
            gl_sin_p  = _pct(gl_sin_n, gl_tot);               gl_con_p  = _pct(gl_con_n, gl_tot);                gl_comp_p = _pct(gl_comp_n, gl_tot)
            _panel_tres(top_gl, "Gobierno Local", gl_sin_n, gl_sin_p, gl_con_n, gl_con_p, gl_comp_n, gl_comp_p, gl_tot)

            fp_tot = agg.get("Indicadores Fuerza Pública", 0)
            fp_sin_n  = agg.get("FP Sin actividades (n)", 0); fp_con_n  = agg.get("FP Con actividades (n)", 0); fp_comp_n = agg.get("FP Completos (n)", 0)
            fp_sin_p  = _pct(fp_sin_n, fp_tot);               fp_con_p  = _pct(fp_con_n, fp_tot);                fp_comp_p = _pct(fp_comp_n, fp_tot)
            _panel_tres(top_fp, "Fuerza Pública", fp_sin_n, fp_sin_p, fp_con_n, fp_con_p, fp_comp_n, fp_comp_p, fp_tot)

            bottom = st.container()
            _resumen_avance(bottom, sin_n, sin_p, con_n, con_p, comp_n, comp_p, total_ind)

            # === NUEVO: filtros y resumen extendido ===
            df_catprob, found_cols = _extra_filters_ui(scope_df, key_prefix="deleg_")
            _render_informativo(df_catprob, found_cols, key_prefix="deleg_")

    # =================== TAB 2: POR DIRECCIÓN REGIONAL ===================
    with tabs[1]:
        st.subheader("Avance por Dirección Regional (DR)")

        drs = sorted(df_dash["DR_inferida"].astype(str).unique().tolist(), key=_dr_sort_key)
        idx_default = next((i for i,v in enumerate(drs) if v and "sin dr" not in v.lower()), 0)

        sel_dr = st.selectbox("Dirección Regional", drs, index=idx_default, key="sel_dr")

        df_dr_sel = df_dash[df_dash["DR_inferida"] == sel_dr]

        scope_df, using_total = _scope_total_o_seleccion(
            df_selected=df_dr_sel,
            df_all_options=df_dash,
            key="toggle_total_dr",
            etiqueta_plural="direcciones regionales"
        )

        if scope_df.empty:
            st.info("No hay registros para esa selección.")
        else:
            agg = scope_df.select_dtypes(include=[np.number]).sum(numeric_only=True)

            total_ind = agg.get("Total Indicadores", np.nan)
            if np.isnan(total_ind):
                total_ind = agg.get("Indicadores Gobierno Local", 0) + agg.get("Indicadores Fuerza Pública", 0)

            sin_n  = agg.get("Sin actividades (n)", 0)
            con_n  = agg.get("Con actividades (n)", 0)
            comp_n = agg.get("Completos (n)", 0)

            def _pct(n, d): 
                return (n / d * 100.0) if d > 0 else 0.0

            sin_p, con_p, comp_p = _pct(sin_n, total_ind), _pct(con_n, total_ind), _pct(comp_n, total_ind)

            titulo_h3 = "Total (todas las DR)" if using_total else sel_dr
            st.markdown(f"<h3 style='text-align:center;margin-top:0;color:#111;'>{titulo_h3}</h3>", unsafe_allow_html=True)

            _render_lineas_block(agg)
            _bar_avance((sin_p, con_p, comp_p), title="Total de indicadores (Gobierno Local + Fuerza Pública)")

            top_gl, top_fp = st.columns(2)
            gl_tot = agg.get("Indicadores Gobierno Local", 0)
            gl_sin_n  = agg.get("GL Sin actividades (n)", 0); gl_con_n  = agg.get("GL Con actividades (n)", 0); gl_comp_n = agg.get("GL Completos (n)", 0)
            gl_sin_p  = _pct(gl_sin_n, gl_tot);               gl_con_p  = _pct(gl_con_n, gl_tot);                gl_comp_p = _pct(gl_comp_n, gl_tot)
            _panel_tres(top_gl, "Gobierno Local", gl_sin_n, gl_sin_p, gl_con_n, gl_con_p, gl_comp_n, gl_comp_p, gl_tot)

            fp_tot = agg.get("Indicadores Fuerza Pública", 0)
            fp_sin_n  = agg.get("FP Sin actividades (n)", 0); fp_con_n  = agg.get("FP Con actividades (n)", 0); fp_comp_n = agg.get("FP Completos (n)", 0)
            fp_sin_p  = _pct(fp_sin_n, fp_tot);               fp_con_p  = _pct(fp_con_n, fp_tot);                fp_comp_p = _pct(fp_comp_n, fp_tot)
            _panel_tres(top_fp, "Fuerza Pública", fp_sin_n, fp_sin_p, fp_con_n, fp_con_p, fp_comp_n, fp_comp_p, fp_tot)

            bottom = st.container()
            _resumen_avance(bottom, sin_n, sin_p, con_n, con_p, comp_n, comp_p, total_ind)

            # === NUEVO: filtros y resumen extendido ===
            df_catprob, found_cols = _extra_filters_ui(scope_df, key_prefix="dr_")
            _render_informativo(df_catprob, found_cols, key_prefix="dr_")

    # =================== TAB 3: SOLO GOBIERNO LOCAL (PROVINCIA) ==========
    with tabs[2]:
        st.subheader("Gobierno Local (filtrar por Provincia)")

        prov_col = _pick_prov_column(df_dash)
        if not prov_col:
            st.warning("No se detectó una columna de **Provincia** en el Excel consolidado. Agrega una columna 'Provincia'.")
        else:
            provincias = sorted(df_dash[prov_col].dropna().astype(str).unique().tolist())
            sel_prov = st.selectbox("Provincia", provincias, index=0, key="sel_prov_only")

            df_prov_sel = df_dash[df_dash[prov_col].astype(str) == sel_prov]

            scope_df, using_total = _scope_total_o_seleccion(
                df_selected=df_prov_sel,
                df_all_options=df_dash,
                key="toggle_total_prov",
                etiqueta_plural="provincias"
            )

            if scope_df.empty:
                st.info("No hay registros para esa selección.")
            else:
                agg = scope_df.select_dtypes(include=[np.number]).sum(numeric_only=True)

                _render_lineas_block(agg)

                def _pct(n, d):
                    return (n / d * 100.0) if d > 0 else 0.0

                gl_tot = agg.get("Indicadores Gobierno Local", 0)
                gl_sin_n  = agg.get("GL Sin actividades (n)", 0)
                gl_con_n  = agg.get("GL Con actividades (n)", 0)
                gl_comp_n = agg.get("GL Completos (n)", 0)

                gl_sin_p = _pct(gl_sin_n, gl_tot)
                gl_con_p = _pct(gl_con_n, gl_tot)
                gl_comp_p = _pct(gl_comp_n, gl_tot)

                titulo_h3 = "Total (todas las provincias)" if using_total else f"Provincia: {sel_prov}"
                st.markdown(
                    f"<h3 style='text-align:center;margin-top:0;color:#111;'>{titulo_h3}</h3>",
                    unsafe_allow_html=True
                )

                _bar_avance((gl_sin_p, gl_con_p, gl_comp_p), title="Total de indicadores (Gobierno Local)")

                _panel_tres(st.container(), "Gobierno Local",
                            gl_sin_n, gl_sin_p, gl_con_n, gl_con_p, gl_comp_n, gl_comp_p, gl_tot)

                if not using_total:
                    delegs = sorted(scope_df["Delegación"].dropna().astype(str).unique().tolist()) if "Delegación" in scope_df.columns else []
                    if delegs:
                        st.markdown(
                            "<div style='margin-top:12px;background:#fff;border:1px solid #e3e3e3;border-radius:8px;padding:12px;'>"
                            f"<div style='font-weight:700;margin-bottom:8px;color:#111;'>Delegaciones en {sel_prov}</div>"
                            + "<ul style='margin:0 0 0 18px;color:#111;'>" +
                            "".join([f"<li>{d}</li>" for d in delegs]) +
                            "</ul></div>",
                            unsafe_allow_html=True
                        )

                # === NUEVO: filtros y resumen extendido ===
                df_catprob, found_cols = _extra_filters_ui(scope_df, key_prefix="prov_")
                _render_informativo(df_catprob, found_cols, key_prefix="prov_")
else:
    st.info("Carga el Excel consolidado para habilitar los dashboards.")
