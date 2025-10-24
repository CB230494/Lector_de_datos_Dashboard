# -*- coding: utf-8 -*-
# ================================================================
# Lector de Matrices (Excel) → Resumen consolidado en Excel + Dashboard
# - Consolidado desde múltiples matrices
# - Dashboard por Delegación / DR / Provincia
# - Bloques de Categorías y Problemáticas (GL / FP / Mixtas) alimentados
#   desde una hoja "detalle" en el MISMO Excel (auto-detectable).
# ================================================================

import io, re, unicodedata
from typing import Dict, Optional, Tuple, List
import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# ------------------------ Utilidades base ----------------------------
def _norm(x: str) -> str:
    if x is None: return ""
    x = str(x)
    x = unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", x).strip().lower()

def _norm_str(s: str) -> str:
    if s is None: return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    s = re.sub(r"[\s_\-]+", " ", s).strip().lower()
    return s

def _to_int(x) -> Optional[int]:
    if x is None: return None
    s = str(x).strip()
    if "%" in s: return None
    digits = re.sub(r"[^\d-]", "", s)
    if not re.fullmatch(r"-?\d{1,3}", digits): return None
    try: return int(digits)
    except: return None

def _to_pct(x) -> Optional[float]:
    if x is None: return None
    s = str(x).replace(",", ".")
    m = re.search(r"-?\d+(\.\d+)?", s)
    if not m: return None
    v = float(m.group())
    if "%" in s or v > 1.0: return max(0.0, min(100.0, v))
    return max(0.0, min(100.0, v * 100.0))

def _read_df(file) -> pd.DataFrame:
    return pd.read_excel(file, engine="openpyxl", header=None, dtype=str)

def _find(df: pd.DataFrame, pattern: str) -> List[Tuple[int,int]]:
    rx = re.compile(pattern)
    out = []
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = df.iat[r,c]
            if val is None: continue
            if rx.search(_norm(val)): out.append((r,c))
    return out

def _neighbors(df: pd.DataFrame, r: int, c: int, up: int, down: int, left: int, right: int):
    r0 = max(0, r - up); r1 = min(df.shape[0]-1, r + down)
    c0 = max(0, c - left); c1 = min(df.shape[1]-1, c + right)
    for i in range(r0, r1+1):
        for j in range(c0, c1+1):
            yield i, j

def _pick_best_count(cands: List[int], max_allowed: int = 60) -> Optional[int]:
    cands = [x for x in cands if x is not None and 0 <= x <= max_allowed]
    return max(cands) if cands else None

# ------------------- Detectores (consolidado) ---------------------
def detect_delegacion(df: pd.DataFrame) -> Optional[str]:
    rx = re.compile(r"^\s*d\d{1,3}\s*[-–]\s*.+\s*$", re.IGNORECASE)
    for r in range(min(15, df.shape[0])):
        for c in range(df.shape[1]):
            raw = df.iat[r,c]
            if raw and rx.match(str(raw)): return str(raw).strip()
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
        if debug and val is not None: st.caption(f"Líneas de Acción detectadas: {val}")
        if val is not None: return val
    la_hits = _find(df, r"\blinea\s+de\s+accion\s*#")
    return len(la_hits) if la_hits else None

# ---- Detección GL/FP y Avance por “lo visible” (consolidado) ----
def detect_indicator_rows(df: pd.DataFrame) -> List[int]:
    rows = []
    for r in range(df.shape[0]):
        left_vals = [df.iat[r, c] for c in range(min(6, df.shape[1]))]
        left_norm = [_norm(v) for v in left_vals]
        if any(v == "gl" for v in left_norm) or any(v == "fp" for v in left_norm):
            rows.append(r)
    return rows

def detect_role_of_row(df: pd.DataFrame, r: int) -> Optional[str]:
    left_vals = [df.iat[r, c] for c in range(min(6, df.shape[1]))]
    left_norm = [_norm(v) for v in left_vals]
    if any(v == "gl" for v in left_norm): return "gl"
    if any(v == "fp" for v in left_norm): return "fp"
    return None

def detect_avance_columns(df: pd.DataFrame) -> List[int]:
    cols = []
    for r in range(df.shape[0]):
        row = [df.iat[r, c] for c in range(df.shape[1])]
        for c, v in enumerate(row):
            if "avance" in _norm(v): cols.append(c)
        if len(cols) >= 2: return sorted(list(set(cols)))
        else: cols = []
    fallback = [10, 15, 20, 25]
    return [c for c in fallback if c < df.shape[1]]

def gl_fp_counts(df: pd.DataFrame) -> Tuple[int, int]:
    rows = detect_indicator_rows(df); gl = fp = 0
    for r in rows:
        role = detect_role_of_row(df, r)
        if role == "gl": gl += 1
        elif role == "fp": fp += 1
    return gl, fp

def _row_status_from_avance(df: pd.DataFrame, r: int, avance_cols: List[int]) -> str:
    vals = [df.iat[r, c] for c in avance_cols]; valsn = [_norm(v) for v in vals]
    if any("complet" in v for v in valsn): return "completos"
    if any("con actividades" in v for v in valsn): return "con_actividades"
    if any("sin actividades" in v for v in valsn): return "sin_actividades"
    return "sin_actividades"

def avance_counts(df: pd.DataFrame) -> Tuple[Dict[str,int], int]:
    rows = detect_indicator_rows(df); avance_cols = detect_avance_columns(df)
    counts = {"completos": 0, "con_actividades": 0, "sin_actividades": 0}
    for r in rows: counts[_row_status_from_avance(df, r, avance_cols)] += 1
    return counts, len(rows)

def avance_counts_by_role(df: pd.DataFrame) -> Tuple[Dict[str,int], int, Dict[str,int], int]:
    rows = detect_indicator_rows(df); avance_cols = detect_avance_columns(df)
    gl_counts = {"completos": 0, "con_actividades": 0, "sin_actividades": 0}
    fp_counts = {"completos": 0, "con_actividades": 0, "sin_actividades": 0}
    n_gl = n_fp = 0
    for r in rows:
        role = detect_role_of_row(df, r); stt = _row_status_from_avance(df, r, avance_cols)
        if role == "gl": gl_counts[stt] += 1; n_gl += 1
        elif role == "fp": fp_counts[stt] += 1; n_fp += 1
    return gl_counts, n_gl, fp_counts, n_fp

def detect_total_indicadores(df: pd.DataFrame) -> Optional[int]:
    hits = _find(df, r"\btotal\s+de\s+indicadores\b")
    for (r,c) in hits:
        cands = []
        for (i,j) in _neighbors(df, r, c, up=0, down=3, left=0, right=6):
            cands.append(_to_int(df.iat[i,j]))
        val = _pick_best_count([x for x in cands if x is not None], max_allowed=120)
        if val is not None: return val
    return None

# --------------------- Proceso de un archivo (consolidado) --------------
def process_file(upload, debug: bool=False) -> Dict:
    df = _read_df(upload)
    deleg = detect_delegacion(df)
    lineas = detect_lineas_accion(df, debug=debug)
    gl, fp = gl_fp_counts(df)

    avance_dict, total_ind = avance_counts(df)
    comp_n = avance_dict["completos"]; con_n = avance_dict["con_actividades"]; sin_n = avance_dict["sin_actividades"]
    pct = lambda n,d: round((n/d)*100.0,1) if d and n is not None else None
    comp_p, con_p, sin_p = pct(comp_n,total_ind), pct(con_n,total_ind), pct(sin_n,total_ind)

    gl_counts, n_gl, fp_counts, n_fp = avance_counts_by_role(df)
    gl_comp_n, gl_con_n, gl_sin_n = gl_counts["completos"], gl_counts["con_actividades"], gl_counts["sin_actividades"]
    gl_comp_p, gl_con_p, gl_sin_p = pct(gl_comp_n,n_gl), pct(gl_con_n,n_gl), pct(gl_sin_n,n_gl)
    fp_comp_n, fp_con_n, fp_sin_n = fp_counts["completos"], fp_counts["con_actividades"], fp_counts["sin_actividades"]
    fp_comp_p, fp_con_p, fp_sin_p = pct(fp_comp_n,n_fp), pct(fp_con_n,n_fp), pct(fp_sin_n,n_fp)

    total_from_label = detect_total_indicadores(df)
    total_out = total_ind if total_ind else total_from_label

    if (gl == 0 or fp == 0) and total_out and gl + fp != total_out:
        if gl == 0 and fp > 0: gl = max(0, total_out - fp)
        elif fp == 0 and gl > 0: fp = max(0, total_out - gl)

    out = {
        "archivo": upload.name,
        "delegacion": deleg,
        "lineas_accion": lineas,
        "completos_n": comp_n, "completos_pct": comp_p,
        "conact_n": con_n, "conact_pct": con_p,
        "sinact_n": sin_n, "sinact_pct": sin_p,
        "indicadores_gl": gl if gl is not None else None,
        "gl_completos_n": gl_comp_n, "gl_completos_pct": gl_comp_p,
        "gl_conact_n": gl_con_n, "gl_conact_pct": gl_con_p,
        "gl_sinact_n": gl_sin_n, "gl_sinact_pct": gl_sin_p,
        "indicadores_fp": fp if fp is not None else None,
        "fp_completos_n": fp_comp_n, "fp_completos_pct": fp_comp_p,
        "fp_conact_n": fp_con_n, "fp_conact_pct": fp_con_p,
        "fp_sinact_n": fp_sin_n, "fp_sinact_pct": fp_sin_p,
        "indicadores_total": total_out if total_out is not None else (gl + fp if (gl or fp) else None),
    }
    if debug:
        st.caption(f"[DEBUG] Avance cols: {detect_avance_columns(df)} | rows GL/FP: {gl}+{fp}={gl+fp}")
    return out

# --------------------------- UI (consolidado) ---------------------------
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
        try: rows.append(process_file(f, debug=debug))
        except Exception as e: failed.append((f.name, str(e)))

    if rows:
        df_out = pd.DataFrame(rows)
        rename = {
            "archivo":"Archivo","delegacion":"Delegación","lineas_accion":"Líneas de Acción",
            "indicadores_gl":"Indicadores Gobierno Local",
            "gl_completos_n":"GL Completos (n)","gl_completos_pct":"GL Completos (%)",
            "gl_conact_n":"GL Con actividades (n)","gl_conact_pct":"GL Con actividades (%)",
            "gl_sinact_n":"GL Sin actividades (n)","gl_sinact_pct":"GL Sin actividades (%)",
            "indicadores_fp":"Indicadores Fuerza Pública",
            "fp_completos_n":"FP Completos (n)","fp_completos_pct":"FP Completos (%)",
            "fp_conact_n":"FP Con actividades (n)","fp_conact_pct":"FP Con actividades (%)",
            "fp_sinact_n":"FP Sin actividades (n)","fp_sinact_pct":"FP Sin actividades (%)",
            "indicadores_total":"Total Indicadores",
            "completos_n":"Completos (n)","completos_pct":"Completos (%)",
            "conact_n":"Con actividades (n)","conact_pct":"Con actividades (%)",
            "sinact_n":"Sin actividades (n)","sinact_pct":"Sin actividades (%)",
        }
        order = [
            "archivo","delegacion","lineas_accion",
            "indicadores_gl","gl_completos_n","gl_completos_pct","gl_conact_n","gl_conact_pct","gl_sinact_n","gl_sinact_pct",
            "indicadores_fp","fp_completos_n","fp_completos_pct","fp_conact_n","fp_conact_pct","fp_sinact_n","fp_sinact_pct",
            "indicadores_total","completos_n","completos_pct","conact_n","conact_pct","sinact_n","sinact_pct",
        ]
        df_out = df_out[order].rename(columns=rename)
        pct_cols = [
            "GL Completos (%)","GL Con actividades (%)","GL Sin actividades (%)",
            "FP Completos (%)","FP Con actividades (%)","FP Sin actividades (%)",
            "Completos (%)","Con actividades (%)","Sin actividades (%)",
        ]
        for col in pct_cols:
            if col in df_out.columns:
                df_out[col] = df_out[col].apply(lambda v: f"{v:.1f}%" if pd.notna(v) else None)
        st.subheader("Resumen previo"); st.dataframe(df_out, use_container_width=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df_out.to_excel(w, index=False, sheet_name="resumen")
        st.download_button("⬇️ Descargar Excel consolidado", data=buf.getvalue(),
                           file_name="resumen_matrices.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if failed:
        st.warning("Algunos archivos no se pudieron procesar automáticamente:")
        for name, err in failed: st.write(f"- {name}: {err}")
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
    2) Opcionalmente, en el mismo archivo, agrega una **hoja “detalle”** con columnas:
       `Delegación`, `Categoría`, `Problemática`, `Lider` (GL/FP/ambos) y opcionalmente `Provincia`.  
    3) Las tarjetas de **Categorías y Problemáticas** usan esa hoja y **respetan los filtros**.
    """)

dash_file = st.file_uploader("Cargar Excel consolidado (mismo archivo con hoja 'detalle')", type=["xlsx"], key="dash_excel")

# -------------------- helpers dashboard / estilos ---------------------
def _to_num_safe(x, pct=False):
    if pd.isna(x): return 0.0
    s = str(x).strip()
    if pct:
        s = s.replace("%", "").replace(",", ".")
        try: return float(s)
        except: return 0.0
    s = s.replace(",", ".")
    try: return float(s)
    except:
        m = re.search(r"-?\d+(\.\d+)?", s)
        return float(m.group()) if m else 0.0

def _big_number(value, label, big_px=72):
    st.markdown(f"""
    <div style="text-align:center;padding:10px;background:#ffffff;border:1px solid #e3e3e3;border-radius:8px;">
        <div style="font-size:14px;color:#666;margin-bottom:6px;">{label}</div>
        <div style="font-size:{big_px}px;font-weight:900;line-height:1;color:#111;">{int(value)}</div>
    </div>""", unsafe_allow_html=True)

COLOR_ROJO="#ED1C24"; COLOR_AMARIL="#F4C542"; COLOR_VERDE="#7AC943"; COLOR_AZUL_H="#1F4E79"

def _bar_avance(pcts_tuple, title=""):
    labels=["Sin actividades","Con actividades","Cumplida"]; values=list(pcts_tuple)
    colors=[COLOR_ROJO,COLOR_AMARIL,COLOR_VERDE]
    fig,ax=plt.subplots(figsize=(5.5,3.5)); fig.patch.set_facecolor("#fff"); ax.set_facecolor("#fff")
    ax.bar(labels, values, color=colors); ax.set_ylim(0,100); ax.set_ylabel("%", color="#111"); ax.set_title(title, color="#111")
    ax.tick_params(axis="x", colors="#111"); ax.tick_params(axis="y", colors="#111")
    for spine in ax.spines.values(): spine.set_color("#999")
    for i,v in enumerate(values): ax.text(i, v+1, f"{v:.0f}%", ha="center", va="bottom", fontsize=10, color="#111")
    st.pyplot(fig, use_container_width=True)

def _panel_tres(col, titulo, n_rojo, p_rojo, n_amar, p_amar, n_verde, p_verde, total):
    with col:
        st.markdown(f"""
        <div style="background:{COLOR_AZUL_H};color:white;padding:10px 12px;border-radius:8px 8px 0 0;
                    font-weight:700;text-align:center;border:1px solid #e3e3e3;border-bottom:0;">{titulo}</div>""", unsafe_allow_html=True)
        c1,c2,c3=st.columns(3)
        c1.markdown(f"""<div style="background:{COLOR_ROJO};color:white;text-align:center;padding:10px;border:1px solid #e3e3e3;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(n_rojo)}</div>
              <div style="font-size:13px;">Sin actividades</div>
              <div style="font-size:16px;font-weight:700;">{p_rojo:.0f}%</div></div>""", unsafe_allow_html=True)
        c2.markdown(f"""<div style="background:{COLOR_AMARIL};color:#111;text-align:center;padding:10px;border:1px solid #e3e3e3;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(n_amar)}</div>
              <div style="font-size:13px;">Con actividades</div>
              <div style="font-size:16px;font-weight:700;">{p_amar:.0f}%</div></div>""", unsafe_allow_html=True)
        c3.markdown(f"""<div style="background:{COLOR_VERDE};color:#111;text-align:center;padding:10px;border:1px solid #e3e3e3;">
              <div style="font-size:36px;font-weight:800;line-height:1;">{int(n_verde)}</div>
              <div style="font-size:13px;">Cumplida</div>
              <div style="font-size:16px;font-weight:700;">{p_verde:.0f}%</div></div>""", unsafe_allow_html=True)
        st.markdown(f"""<div style="text-align:center;border:1px solid #e3e3e3;border-top:0;padding:10px;border-radius:0 0 8px 8px;background:#ffffff;color:#111;">
            <div style="font-size:13px;color:#666;margin-bottom:4px;">Total de indicadores</div>
            <div style="font-size:44px;font-weight:900;line-height:1;">{int(total)}</div></div>""", unsafe_allow_html=True)

def _ensure_numeric(df):
    cols_n = ["GL Completos (n)","GL Con actividades (n)","GL Sin actividades (n)",
              "FP Completos (n)","FP Con actividades (n)","FP Sin actividades (n)",
              "Completos (n)","Con actividades (n)","Sin actividades (n)",
              "Indicadores Gobierno Local","Indicadores Fuerza Pública","Total Indicadores",
              "Líneas de Acción","Líneas de Acción Gobierno Local","Líneas de Acción Fuerza Pública","Líneas de Acción Mixtas"]
    cols_p = ["GL Completos (%)","GL Con actividades (%)","GL Sin actividades (%)",
              "FP Completos (%)","FP Con actividades (%)","FP Sin actividades (%)",
              "Completos (%)","Con actividades (%)","Sin actividades (%)"]
    for c in cols_n:
        if c in df.columns: df[c] = df[c].apply(_to_num_safe)
    for c in cols_p:
        if c in df.columns: df[c] = df[c].apply(lambda v: _to_num_safe(v, pct=True))
    return df

# ---------- DR / Provincia: detección (y fallback por Delegación) -------
def _infer_dr_from_delegacion(name: str) -> str:
    if not isinstance(name, str): return "Sin DR / No identificado"
    m = re.search(r"(R\s*\d+)", name, flags=re.IGNORECASE)
    if m: return m.group(1).upper().replace(" ", "")
    m = re.search(r"(DR-\s*\d+\S*)", name, flags=re.IGNORECASE)
    if m: return m.group(1).replace(" ", "")
    return "Sin DR / No identificado"

def _pick_dr_column(df: pd.DataFrame) -> Optional[str]:
    norm_map = {_norm_str(c): c for c in df.columns}
    for cand in ["direccionregional","direccion regional","dirregional","dr","region","regional","direccion","direccionreg","regiones"]:
        if cand in norm_map: return norm_map[cand]
    for k, real in norm_map.items():
        if ("direccion" in k and "regional" in k) or k=="dr": return real
    return None

def _dr_sort_key(s: str):
    if not isinstance(s, str): return (999, "")
    m = re.search(r"r\s*([0-9]+)", s, flags=re.IGNORECASE)
    num = int(m.group(1)) if m else 999
    return (num, s)

def _pick_prov_column(df: pd.DataFrame) -> Optional[str]:
    norm_map = {_norm_str(c): c for c in df.columns}
    candidates = ["provincia","province","prov","provincial","provincia nombre","nom provincia"]
    for cand in candidates:
        if cand in norm_map: return norm_map[cand]
    for k, real in norm_map.items():
        if "provinc" in k: return real
    return None

# --------- helpers “detalle” (Categoría/Problemática/Lider) -------------
def _pick_col(df: pd.DataFrame, candidates_exact: List[str], substrs: List[str]) -> Optional[str]:
    if df is None or df.empty: return None
    norm_map = {_norm_str(c): c for c in df.columns}
    for c in candidates_exact:
        if c in norm_map: return norm_map[c]
    for k, real in norm_map.items():
        if any(s in k for s in substrs): return real
    return None

def _pick_categoria_col(df):   return _pick_col(df,
    ["categoria","categoría","eje","linea","línea"], ["categori","eje","linea","línea"])
def _pick_problematica_col(df):return _pick_col(df,
    ["problematica","problemática","problema","descriptor","tema"], ["problem","descrip","tema"])
def _pick_rol_col(df):         return _pick_col(df,
    ["lider","líder","rol","tipo","responsable","encargado","actor"],
    ["lider","líder","rol","tipo","respons","encarg","actor","gl","fp","mixt","ambos"])
def _pick_deleg_col(df):       return _pick_col(df,
    ["delegacion","delegación"], ["deleg"])
def _pick_prov_det_col(df):    return _pick_prov_column(df)  # reusa lógica

def _map_rol_value(x: str) -> str:
    s = _norm_str(x)
    if any(k in s for k in ["mixt","ambos","compartid"]): return "Mixtas"
    if s in ("gl","g l") or "gobierno local" in s or "municip" in s or "alcald" in s: return "Gobierno Local"
    if s in ("fp","f p") or "fuerza publica" in s or "fuerza pública" in s or "policia" in s or "policía" in s: return "Fuerza Pública"
    return "Gobierno Local"

def _aggregate_roles(df_scope: pd.DataFrame, target_col: str, rol_col: str) -> pd.DataFrame:
    if df_scope.empty: return pd.DataFrame()
    tmp = df_scope[[target_col, rol_col]].dropna().copy()
    if tmp.empty: return pd.DataFrame()
    tmp["__ROL__"] = tmp[rol_col].astype(str).apply(_map_rol_value); tmp["__ONE__"] = 1
    agg = tmp.groupby([target_col, "__ROL__"])["__ONE__"].sum().reset_index()
    pivot = agg.pivot(index=target_col, columns="__ROL__", values="__ONE__").fillna(0).astype(int)
    for col in ["Gobierno Local","Fuerza Pública","Mixtas"]:
        if col not in pivot.columns: pivot[col]=0
    pivot["Total"] = pivot["Gobierno Local"]+pivot["Fuerza Pública"]+pivot["Mixtas"]
    pivot = pivot.sort_values(["Total","Gobierno Local","Fuerza Pública","Mixtas"], ascending=False)
    return pivot

def _pill(text, value, bg, fg="#111"):
    return f"""<span style="display:inline-block;padding:4px 8px;margin-right:6px;border-radius:999px;
               background:{bg};color:{fg};font-size:12px;border:1px solid #e3e3e3;">{text}: <b>{int(value)}</b></span>"""

def _render_cards_from_pivot(pivot: pd.DataFrame, title: str, max_inline: int = 24):
    st.markdown(f"<h4 style='margin:12px 0 8px;color:#111;'>{title}</h4>", unsafe_allow_html=True)
    if pivot.empty:
        st.info("No hay datos suficientes para este bloque (se requieren columnas de Delegación, Lider y Categoría/Problemática).")
        return

    records = pivot.reset_index().to_dict("records")
    head, tail = records[:max_inline], records[max_inline:]

    def _name(rec): return str(rec[pivot.index.name])
    def _val(rec, key): return int(rec.get(key, 0) or 0)
    def _card_html(name, gl, fp, mx, total):
        return f"""
        <div style="background:#fff;border:1px solid #eaeaea;border-radius:10px;padding:12px;margin-bottom:10px;">
          <div style="font-size:15px;font-weight:700;color:#111;margin-bottom:8px;word-break:break-word;">{name}</div>
          <div>{_pill("GL", gl, "#F0F9F0")}{_pill("FP", fp, "#F0F4FF")}{_pill("Mixtas", mx, "#FFF8E6")}{_pill("Total", total, "#F7F7F7")}</div>
        </div>"""

    def _render_list(items):
        cols = st.columns(3, gap="small")
        for i, rec in enumerate(items):
            name = _name(rec)
            gl = _val(rec, "Gobierno Local"); fp = _val(rec, "Fuerza Pública"); mx = _val(rec, "Mixtas")
            total = _val(rec, "Total") if "Total" in rec else gl+fp+mx
            with cols[i % 3]: st.markdown(_card_html(name, gl, fp, mx, total), unsafe_allow_html=True)

    _render_list(head)
    if tail:
        with st.expander(f"Ver más ({len(tail)} adicionales)"): _render_list(tail)

# --------------- Overrides para columnas/hoja “detalle” ----------------
def _string_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if df[c].dtype == object]

def _set_override_default(name: str, value: Optional[str]):
    if name not in st.session_state: st.session_state[name] = value

def _resolve_col(df: pd.DataFrame, auto_pick_fn, override_name: str) -> Optional[str]:
    ov = st.session_state.get(override_name)
    if ov and ov != "(auto)" and ov in df.columns: return ov
    return auto_pick_fn(df)

def _auto_pick_detail_sheet(xls: pd.ExcelFile) -> Optional[str]:
    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet)
            cols = [c.lower() for c in df.columns.astype(str)]
            if any("deleg" in c for c in cols) and any("ategor" in c for c in cols) and any("roblem" in c for c in cols):
                return sheet
        except: pass
    return None

# ============================= MAIN DASHBOARD =============================
if dash_file:
    # --- Cargar Excel (todas las hojas disponibles)
    xls = pd.ExcelFile(dash_file)
    # Consolidado (resumen)
    try:
        df_dash = pd.read_excel(dash_file, sheet_name="resumen")
    except Exception:
        # usa primera hoja con columnas del consolidado si no existe "resumen"
        df_dash = xls.parse(xls.sheet_names[0])
    df_dash = _ensure_numeric(df_dash.copy())

    # ======= Detección de hoja “detalle” (Categoría/Problemática) =======
    auto_detail_sheet = _auto_pick_detail_sheet(xls)
    with st.expander("🛠️ Detección avanzada (opcional)", expanded=False):
        st.caption("Si no se detecta automáticamente tu hoja de detalle/columnas, selecciónalas aquí.")
        st.write(f"**Hoja de detalle detectada:** `{auto_detail_sheet}`")
        detail_sheet = st.selectbox("Hoja con detalle (Categoría / Problemática / Lider / Delegación)",
                                    ["(auto)"] + xls.sheet_names,
                                    index=(["(auto)"] + xls.sheet_names).index(auto_detail_sheet) if auto_detail_sheet in xls.sheet_names else 0,
                                    key="ov_detail_sheet")
    if detail_sheet and detail_sheet != "(auto)":
        df_detalle = xls.parse(detail_sheet)
    else:
        df_detalle = xls.parse(auto_detail_sheet) if auto_detail_sheet else pd.DataFrame()

    # ---- Overrides de columnas del DETALLE ----
    with st.expander("🧭 Columnas del detalle (overrides)", expanded=False):
        cols_list = list(df_detalle.columns.astype(str)) if not df_detalle.empty else []
        auto_cat  = _pick_categoria_col(df_detalle) if not df_detalle.empty else None
        auto_prob = _pick_problematica_col(df_detalle) if not df_detalle.empty else None
        auto_rol  = _pick_rol_col(df_detalle) if not df_detalle.empty else None
        auto_deleg= _pick_deleg_col(df_detalle) if not df_detalle.empty else None
        auto_prov = _pick_prov_det_col(df_detalle) if not df_detalle.empty else None

        _set_override_default("ov_col_categoria", auto_cat)
        _set_override_default("ov_col_problematica", auto_prob)
        _set_override_default("ov_col_rol", auto_rol)
        _set_override_default("ov_col_deleg", auto_deleg)
        _set_override_default("ov_col_prov", auto_prov)

        st.write("**Detectado automáticamente en el detalle:**")
        st.write(f"- Categoría: `{auto_cat}`  ·  Problemática: `{auto_prob}`  ·  Lider/Rol: `{auto_rol}`")
        st.write(f"- Delegación: `{auto_deleg}`  ·  Provincia: `{auto_prov}`")

        if cols_list:
            st.session_state["ov_col_categoria"] = st.selectbox("Columna de Categoría", ["(auto)"] + cols_list,
                                                                index=(["(auto)"] + cols_list).index(auto_cat) if auto_cat in cols_list else 0)
            st.session_state["ov_col_problematica"] = st.selectbox("Columna de Problemática", ["(auto)"] + cols_list,
                                                                   index=(["(auto)"] + cols_list).index(auto_prob) if auto_prob in cols_list else 0)
            st.session_state["ov_col_rol"] = st.selectbox("Columna de Lider/Rol/Tipo (GL/FP/Mixta)", ["(auto)"] + cols_list,
                                                          index=(["(auto)"] + cols_list).index(auto_rol) if auto_rol in cols_list else 0)
            st.session_state["ov_col_deleg"] = st.selectbox("Columna de Delegación (detalle)", ["(auto)"] + cols_list,
                                                            index=(["(auto)"] + cols_list).index(auto_deleg) if auto_deleg in cols_list else 0)
            st.session_state["ov_col_prov"] = st.selectbox("Columna de Provincia (detalle, opcional)", ["(auto)"] + ["(ninguna)"] + cols_list,
                                                           index=(["(auto)"] + ["(ninguna)"] + cols_list).index(auto_prov) if auto_prov in cols_list else 0)

    # ----- Añadir DR_inferida al CONSOLIDADO -----
    dr_col = _pick_dr_column(df_dash)
    if dr_col:
        tmp = (df_dash[dr_col].astype(str).apply(_norm_str).str.replace(r"\s+", " ", regex=True).str.strip())
        tmp = tmp.str.replace(r"(^r\s*\d+)\s*", lambda m: m.group(1).upper().replace(" ","")+" ", regex=True)
        tmp = tmp.replace({"": "Sin DR / No identificado", "nan": "Sin DR / No identificado", "none": "Sin DR / No identificado"})
        df_dash["DR_inferida"] = tmp
    else:
        df_dash["DR_inferida"] = df_dash.get("Delegación","").apply(_infer_dr_from_delegacion)

    # ----- Preparar DETALLE normalizado (Delegación / Provincia / DR) -----
    if not df_detalle.empty:
        dcol = _resolve_col(df_detalle, _pick_deleg_col, "ov_col_deleg")
        pcol = None
        ov = st.session_state.get("ov_col_prov")
        if ov and ov not in ("(auto)","(ninguna)") and ov in df_detalle.columns: pcol = ov
        else: pcol = _pick_prov_det_col(df_detalle)

        df_detalle["_Delegación"] = df_detalle[dcol].astype(str) if dcol else ""
        df_detalle["_DR_inferida"] = df_detalle["_Delegación"].apply(_infer_dr_from_delegacion)
        if pcol: df_detalle["_Provincia"] = df_detalle[pcol].astype(str)
        else:    df_detalle["_Provincia"] = np.nan

    # ------------------- Pestañas -------------------
    tabs = st.tabs(["🏢 Por Delegación", "🗺️ Por Dirección Regional", "🏛️ Gobierno Local (por Provincia)"])

    # =============== helpers de render de “detalle” =================
    def _render_detalle_blocks(df_scope_detalle: pd.DataFrame, titulo_scope: str):
        if df_scope_detalle is None or df_scope_detalle.empty:
            st.info("No hay filas de detalle para esta selección.")
            return
        cat_col = _resolve_col(df_detalle, _pick_categoria_col, "ov_col_categoria")
        prob_col= _resolve_col(df_detalle, _pick_problematica_col, "ov_col_problematica")
        rol_col = _resolve_col(df_detalle, _pick_rol_col, "ov_col_rol")

        st.markdown("<hr/>", unsafe_allow_html=True)
        st.markdown("<h3 style='margin:0;color:#111;'>📦 Categorías y Problemáticas (por rol)</h3>", unsafe_allow_html=True)
        st.caption(f"Basado en el subconjunto activo del filtro: **{titulo_scope}**.")
        # Limpieza básica
        def cln(s): return s.astype(str).str.replace("\u00A0"," ",regex=False).str.strip()
        for c in [cat_col, prob_col, rol_col]:
            if c and c in df_scope_detalle.columns: df_scope_detalle[c] = cln(df_scope_detalle[c])

        if rol_col and cat_col:
            pv_cat = _aggregate_roles(df_scope_detalle, target_col=cat_col, rol_col=rol_col)
            pv_cat.index.name = "Categoría"
            _render_cards_from_pivot(pv_cat, "Categorías (GL / FP / Mixtas)")
        else:
            st.info("Falta columna de **Lider/Rol/Tipo** o de **Categoría** para este bloque.")

        if rol_col and prob_col:
            pv_prob = _aggregate_roles(df_scope_detalle, target_col=prob_col, rol_col=rol_col)
            pv_prob.index.name = "Problemática"
            _render_cards_from_pivot(pv_prob, "Problemáticas (GL / FP / Mixtas)")
        else:
            st.info("Falta columna de **Lider/Rol/Tipo** o de **Problemática** para este bloque.")

    # ======================= TAB 1: POR DELEGACIÓN =======================
    with tabs[0]:
        st.subheader("Avance por Delegación Policial")

        if "Delegación" not in df_dash.columns:
            st.info("El Excel no contiene la columna 'Delegación'.")
        else:
            delegs = sorted(df_dash["Delegación"].dropna().astype(str).unique().tolist())
            sel = st.selectbox("Delegación Policial", delegs, index=0, key="sel_deleg")

            dsel = df_dash[df_dash["Delegación"] == sel]

            # Toggle total/selección
            use_total = st.toggle("Habilitar mostrar total de datos", value=False, key="toggle_total_deleg")
            scope_df = df_dash if use_total else dsel

            agg = scope_df.select_dtypes(include=[np.number]).sum(numeric_only=True)
            total_ind = agg.get("Total Indicadores", np.nan)
            if np.isnan(total_ind):
                total_ind = agg.get("Indicadores Gobierno Local", 0) + agg.get("Indicadores Fuerza Pública", 0)
            sin_n  = agg.get("Sin actividades (n)", 0); con_n  = agg.get("Con actividades (n)", 0); comp_n = agg.get("Completos (n)", 0)
            pct = lambda n,d: (n/d*100.0) if d>0 else 0.0
            sin_p, con_p, comp_p = pct(sin_n,total_ind), pct(con_n,total_ind), pct(comp_n,total_ind)
            titulo_h3 = "Total (todas las delegaciones)" if use_total else sel
            st.markdown(f"<h3 style='text-align:center;margin-top:0;color:#111;'>{titulo_h3}</h3>", unsafe_allow_html=True)
            # Bloques de números + barras
            total, gl, fp, mx = total_ind, agg.get("Indicadores Gobierno Local", 0), agg.get("Indicadores Fuerza Pública", 0), 0
            _big_number(total, "Líneas de Acción")
            _bar_avance((sin_p, con_p, comp_p), title="Total de indicadores (Gobierno Local + Fuerza Pública)")

            top_gl, top_fp = st.columns(2)
            gl_tot = agg.get("Indicadores Gobierno Local", 0)
            gl_sin_n  = agg.get("GL Sin actividades (n)", 0); gl_con_n  = agg.get("GL Con actividades (n)", 0); gl_comp_n = agg.get("GL Completos (n)", 0)
            gl_sin_p  = pct(gl_sin_n, gl_tot); gl_con_p = pct(gl_con_n, gl_tot); gl_comp_p = pct(gl_comp_n, gl_tot)
            _panel_tres(top_gl, "Gobierno Local", gl_sin_n, gl_sin_p, gl_con_n, gl_con_p, gl_comp_n, gl_comp_p, gl_tot)

            fp_tot = agg.get("Indicadores Fuerza Pública", 0)
            fp_sin_n  = agg.get("FP Sin actividades (n)", 0); fp_con_n  = agg.get("FP Con actividades (n)", 0); fp_comp_n = agg.get("FP Completos (n)", 0)
            fp_sin_p  = pct(fp_sin_n, fp_tot); fp_con_p = pct(fp_con_n, fp_tot); fp_comp_p = pct(fp_comp_n, fp_tot)
            _panel_tres(top_fp, "Fuerza Pública", fp_sin_n, fp_sin_p, fp_con_n, fp_con_p, fp_comp_n, fp_comp_p, fp_tot)

            # ======== DETALLE: filtrar por delegaciones visibles =========
            if not df_detalle.empty:
                vis_delegs = (scope_df["Delegación"].dropna().astype(str).unique().tolist()
                              if "Delegación" in scope_df.columns else [])
                dcol = _resolve_col(df_detalle, _pick_deleg_col, "ov_col_deleg")
                if dcol:
                    scope_det = df_detalle[df_detalle[dcol].astype(str).isin(vis_delegs)]
                    _render_detalle_blocks(scope_det.copy(), titulo_h3)
                else:
                    st.info("La hoja de detalle no tiene columna de Delegación.")
            else:
                st.info("No se cargó hoja de detalle con Categoría/Problemática.")

    # =================== TAB 2: POR DIRECCIÓN REGIONAL ===================
    with tabs[1]:
        st.subheader("Avance por Dirección Regional (DR)")

        drs = sorted(df_dash["DR_inferida"].astype(str).unique().tolist(), key=_dr_sort_key)
        idx_default = next((i for i,v in enumerate(drs) if v and "sin dr" not in v.lower()), 0)
        sel_dr = st.selectbox("Dirección Regional", drs, index=idx_default, key="sel_dr")

        df_dr_sel = df_dash[df_dash["DR_inferida"] == sel_dr]
        use_total = st.toggle("Habilitar mostrar total de datos", value=False, key="toggle_total_dr")
        scope_df = df_dash if use_total else df_dr_sel

        if scope_df.empty:
            st.info("No hay registros para esa selección.")
        else:
            agg = scope_df.select_dtypes(include=[np.number]).sum(numeric_only=True)
            total_ind = agg.get("Total Indicadores", np.nan)
            if np.isnan(total_ind):
                total_ind = agg.get("Indicadores Gobierno Local", 0) + agg.get("Indicadores Fuerza Pública", 0)
            sin_n  = agg.get("Sin actividades (n)", 0); con_n  = agg.get("Con actividades (n)", 0); comp_n = agg.get("Completos (n)", 0)
            pct = lambda n,d: (n/d*100.0) if d>0 else 0.0
            sin_p, con_p, comp_p = pct(sin_n,total_ind), pct(con_n,total_ind), pct(comp_n,total_ind)
            titulo_h3 = "Total (todas las DR)" if use_total else sel_dr
            st.markdown(f"<h3 style='text-align:center;margin-top:0;color:#111;'>{titulo_h3}</h3>", unsafe_allow_html=True)

            _big_number(total_ind, "Líneas de Acción")
            _bar_avance((sin_p, con_p, comp_p), title="Total de indicadores (Gobierno Local + Fuerza Pública)")

            top_gl, top_fp = st.columns(2)
            gl_tot = agg.get("Indicadores Gobierno Local", 0)
            gl_sin_n  = agg.get("GL Sin actividades (n)", 0); gl_con_n  = agg.get("GL Con actividades (n)", 0); gl_comp_n = agg.get("GL Completos (n)", 0)
            gl_sin_p  = pct(gl_sin_n, gl_tot); gl_con_p = pct(gl_con_n, gl_tot); gl_comp_p = pct(gl_comp_n, gl_tot)
            _panel_tres(top_gl, "Gobierno Local", gl_sin_n, gl_sin_p, gl_con_n, gl_con_p, gl_comp_n, gl_comp_p, gl_tot)

            fp_tot = agg.get("Indicadores Fuerza Pública", 0)
            fp_sin_n  = agg.get("FP Sin actividades (n)", 0); fp_con_n  = agg.get("FP Con actividades (n)", 0); fp_comp_n = agg.get("FP Completos (n)", 0)
            fp_sin_p  = pct(fp_sin_n, fp_tot); fp_con_p = pct(fp_con_n, fp_tot); fp_comp_p = pct(fp_comp_n, fp_tot)
            _panel_tres(top_fp, "Fuerza Pública", fp_sin_n, fp_sin_p, fp_con_n, fp_con_p, fp_comp_n, fp_comp_p, fp_tot)

            # ======== DETALLE: usar Delegación→DR_inferida del DETALLE ========
            if not df_detalle.empty:
                dcol = _resolve_col(df_detalle, _pick_deleg_col, "ov_col_deleg")
                if dcol:
                    # map delegaciones del detalle cuya DR coincide
                    tmp = df_detalle.copy()
                    tmp["_DR_inferida"] = tmp["_DR_inferida"] if "_DR_inferida" in tmp.columns else tmp[dcol].astype(str).apply(_infer_dr_from_delegacion)
                    scope_det = tmp[tmp["_DR_inferida"] == (sel_dr if not use_total else tmp["_DR_inferida"])]
                    if use_total: scope_det = tmp  # todas las DR
                    _render_detalle_blocks(scope_det.copy(), titulo_h3)
                else:
                    st.info("La hoja de detalle no tiene columna de Delegación.")
            else:
                st.info("No se cargó hoja de detalle con Categoría/Problemática.")

    # =================== TAB 3: SOLO GOBIERNO LOCAL (PROVINCIA) ==========
    with tabs[2]:
        st.subheader("Gobierno Local (filtrar por Provincia)")

        prov_col_dash = _pick_prov_column(df_dash)
        if not prov_col_dash and (df_detalle.empty or "_Provincia" not in df_detalle.columns):
            st.warning("No se detectó una columna de **Provincia** ni en el consolidado ni en el detalle. Añade una columna 'Provincia' en el detalle para habilitar este filtro.")
        else:
            # Prioriza provincia del DETALLE (más granular)
            prov_series = (df_detalle["_Provincia"].dropna().astype(str) if not df_detalle.empty and "_Provincia" in df_detalle.columns
                           else df_dash[prov_col_dash].dropna().astype(str))
            provincias = sorted(prov_series.unique().tolist())
            if not provincias:
                st.info("No hay provincias disponibles.")
            else:
                sel_prov = st.selectbox("Provincia", provincias, index=0, key="sel_prov_only")

                # Consolidado (para números de avance)
                if prov_col_dash and prov_col_dash in df_dash.columns:
                    df_prov_sel = df_dash[df_dash[prov_col_dash].astype(str) == sel_prov]
                else:
                    # Si solo existe provincia en DETALLE, tomamos delegaciones de esa provincia para armar scope en números
                    dcol = _resolve_col(df_detalle, _pick_deleg_col, "ov_col_deleg")
                    vis_delegs = df_detalle[df_detalle["_Provincia"].astype(str)==sel_prov][dcol].dropna().astype(str).unique().tolist() if not df_detalle.empty and dcol else []
                    df_prov_sel = df_dash[df_dash["Delegación"].astype(str).isin(vis_delegs)] if "Delegación" in df_dash.columns else df_dash.iloc[0:0]

                use_total = st.toggle("Habilitar mostrar total de datos", value=False, key="toggle_total_prov")
                scope_df = df_dash if use_total else df_prov_sel

                if scope_df.empty:
                    st.info("No hay registros para esa selección.")
                else:
                    agg = scope_df.select_dtypes(include=[np.number]).sum(numeric_only=True)
                    gl_tot = agg.get("Indicadores Gobierno Local", 0)
                    gl_sin_n = agg.get("GL Sin actividades (n)", 0); gl_con_n = agg.get("GL Con actividades (n)", 0); gl_comp_n = agg.get("GL Completos (n)", 0)
                    pct = lambda n,d: (n/d*100.0) if d>0 else 0.0
                    gl_sin_p, gl_con_p, gl_comp_p = pct(gl_sin_n,gl_tot), pct(gl_con_n,gl_tot), pct(gl_comp_n,gl_tot)
                    titulo_h3 = "Total (todas las provincias)" if use_total else f"Provincia: {sel_prov}"
                    st.markdown(f"<h3 style='text-align:center;margin-top:0;color:#111;'>{titulo_h3}</h3>", unsafe_allow_html=True)
                    _bar_avance((gl_sin_p, gl_con_p, gl_comp_p), title="Total de indicadores (Gobierno Local)")
                    _panel_tres(st.container(), "Gobierno Local", gl_sin_n, gl_sin_p, gl_con_n, gl_con_p, gl_comp_n, gl_comp_p, gl_tot)

                    # ======== DETALLE por provincia ========
                    if not df_detalle.empty and "_Provincia" in df_detalle.columns:
                        dcol = _resolve_col(df_detalle, _pick_deleg_col, "ov_col_deleg")
                        scope_det = df_detalle[df_detalle["_Provincia"].astype(str)==sel_prov] if not use_total else df_detalle
                        _render_detalle_blocks(scope_det.copy(), titulo_h3)
                    else:
                        st.info("El detalle no tiene columna de Provincia.")
else:
    st.info("Carga el Excel consolidado para habilitar los dashboards.")


                # Categorías y Problemáticas (por rol)
                _render_categorias_y_problemas(scope_df, place_after=titulo_h3)
else:
    st.info("Carga el Excel consolidado para habilitar los dashboards.")

