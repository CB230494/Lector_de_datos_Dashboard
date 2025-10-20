# -*- coding: utf-8 -*-
# ================================================================
# Lector de Matrices (Excel) → Resumen consolidado en Excel
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
    if x is None: return None
    m = re.search(r"-?\d+", str(x).replace(",", ""))
    return int(m.group()) if m else None

def _to_pct(x) -> Optional[float]:
    if x is None: return None
    s = str(x).replace(",", ".")
    m = re.search(r"-?\d+(\.\d+)?", s)
    if not m: return None
    v = float(m.group())
    # si tiene símbolo % o es >1, ya está en 0–100
    if "%" in s or v > 1.0: 
        return v
    return v * 100.0

def _read_df(file) -> pd.DataFrame:
    return pd.read_excel(file, engine="openpyxl", header=None, dtype=str)

def _find(df: pd.DataFrame, pattern: str) -> List[Tuple[int,int]]:
    rx = re.compile(pattern)
    out = []
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            if df.iat[r,c] is None: 
                continue
            if rx.search(_norm(df.iat[r,c])):
                out.append((r,c))
    return out

def _neighbors(df: pd.DataFrame, r: int, c: int, down: int, right: int) -> List[Tuple[int,int]]:
    R = min(df.shape[0], r+down+1)
    C = min(df.shape[1], c+right+1)
    return [(i,j) for i in range(r, R) for j in range(c, C)]

# ------------------- Detectores robustos ------------------------
def detect_delegacion(df: pd.DataFrame) -> Optional[str]:
    # Ej: "D35-Orotina"
    for r in range(min(12, df.shape[0])):
        for c in range(df.shape[1]):
            raw = df.iat[r,c]
            if not raw: continue
            if re.match(r"(?i)^\s*d\d{1,3}\s*[-–]\s*.+\s*$", str(raw)):
                return str(raw).strip()
    # Fallback por rótulo "delegación"
    hits = _find(df, r"\bdelegaci[oó]n\b")
    for (r,c) in hits:
        v = df.iat[r, c+1] if c+1 < df.shape[1] else None
        if v: return str(v).strip()
    return None

def detect_lineas_accion(df: pd.DataFrame) -> Optional[int]:
    # Busca el rótulo y toma el **máximo** entero en un vecindario pequeño para evitar “1” residuales
    hits = _find(df, r"lineas?\s*de\s*accion")
    for (r,c) in hits:
        cand = []
        for (i,j) in _neighbors(df, r, max(0,c-1), down=6, right=3):
            cand.append(_to_int(df.iat[i,j]))
        cand = [x for x in cand if x is not None]
        if cand: 
            return max(cand)
    return None

def _row_has_all(row_vals: List[str], needed: List[str]) -> bool:
    normed = [_norm(v) for v in row_vals]
    return all(any(k in cell for cell in normed) for k in needed)

def detect_avance_indicadores(df: pd.DataFrame) -> Dict[str, Dict[str, Optional[float]]]:
    """
    Devuelve:
      {'completos': {'n','%'}, 'con_actividades': {'n','%'}, 'sin_actividades': {'n','%'}}
    """
    res = {
        "completos": {"n": None, "%": None},
        "con_actividades": {"n": None, "%": None},
        "sin_actividades": {"n": None, "%": None},
    }

    # 1) Encuentra cualquier fila que contenga los 3 encabezados
    for r in range(df.shape[0]):
        row = [str(x) if x is not None else "" for x in df.iloc[r,:].tolist()]
        if _row_has_all(row, ["complet", "con actividades", "sin actividades"]):
            hdr = [str(x) if x is not None else "" for x in df.iloc[r,:].tolist()]
            nums = [str(x) if x is not None else "" for x in df.iloc[r+1,:].tolist()] if r+1<df.shape[0] else []
            pcts = [str(x) if x is not None else "" for x in df.iloc[r+2,:].tolist()] if r+2<df.shape[0] else []

            def pick(key_sub: str) -> Tuple[Optional[int], Optional[float]]:
                idx = None
                for j,h in enumerate(hdr):
                    if key_sub in _norm(h):
                        idx = j; break
                if idx is None: return None, None
                n = _to_int(nums[idx]) if idx < len(nums) else None
                p = _to_pct(pcts[idx]) if idx < len(pcts) else None
                # si vienen mezclados
                if n is None and idx < len(nums):
                    m = _to_pct(nums[idx])
                    if m is not None: p = m
                if p is None and idx < len(pcts):
                    m = _to_int(pcts[idx])
                    if m is not None: n = m
                # normaliza ceros “0%”
                if p is None and idx < len(pcts) and str(pcts[idx]).strip()=="0":
                    p = 0.0
                return n, p

            n,p = pick("complet");           res["completos"] = {"n": n, "%": p}
            n,p = pick("con actividades");   res["con_actividades"] = {"n": n, "%": p}
            n,p = pick("sin actividades");   res["sin_actividades"] = {"n": n, "%": p}
            return res
    return res

def _nearest_category_anchor(df: pd.DataFrame, categoria: str) -> List[Tuple[int,int]]:
    # Vínculos cercanos al texto “Gobierno Local” o “Fuerza Pública”
    return _find(df, _norm(categoria))

def detect_indicadores_categoria(df: pd.DataFrame, categoria: str) -> Optional[int]:
    """
    Busca el número grande del recuadro que dice “Indicadores” dentro del bloque
    de la categoría (Gobierno Local / Fuerza Pública). 
    - Encuentra “Indicadores”
    - Verifica que cerca exista el rótulo de la categoría
    - Toma el entero situado 1–3 filas arriba (misma columna o ±1 col)
    """
    # anclas de categoría
    anchors = _nearest_category_anchor(df, categoria)
    if not anchors:
        # si no hallamos ancla, intentamos igualmente por “indicadores” en todo el sheet
        anchors = [(0,0)]

    ind_cells = _find(df, r"\bindicadores\b")
    best = None
    for (ri,ci) in ind_cells:
        # ¿hay una ancla de categoría cerca?
        if anchors != [(0,0)]:
            if not any(abs(ri-ra)<=12 and abs(ci-ca)<=12 for (ra,ca) in anchors):
                continue
        # busca el número justo encima (misma o ±1 columna) a 1–3 filas
        vals = []
        for up in (1,2,3):
            rr = ri - up
            if rr < 0: break
            for dc in (-1,0,1):
                cc = ci + dc
                if 0<=cc<df.shape[1]:
                    vals.append(_to_int(df.iat[rr,cc]))
        vals = [v for v in vals if v is not None]
        cand = max(vals) if vals else None
        if cand is not None:
            best = cand if best is None else max(best, cand)
    return best

# --------------------- Proceso de un archivo --------------------
def process_file(upload) -> Dict:
    df = _read_df(upload)
    out = {
        "archivo": upload.name,
        "delegacion": detect_delegacion(df),
        "lineas_accion": detect_lineas_accion(df),

        "completos_n": None, "completos_pct": None,
        "conact_n": None, "conact_pct": None,
        "sinact_n": None, "sinact_pct": None,

        "indicadores_gl": detect_indicadores_categoria(df, "gobierno local"),
        "indicadores_fp": detect_indicadores_categoria(df, "fuerza publica"),
    }
    av = detect_avance_indicadores(df)
    out["completos_n"]  = av["completos"]["n"]
    out["completos_pct"]= av["completos"]["%"]
    out["conact_n"]     = av["con_actividades"]["n"]
    out["conact_pct"]   = av["con_actividades"]["%"]
    out["sinact_n"]     = av["sin_actividades"]["n"]
    out["sinact_pct"]   = av["sin_actividades"]["%"]

    gl = out["indicadores_gl"] or 0
    fp = out["indicadores_fp"] or 0
    out["indicadores_total"] = (gl + fp) if (gl or fp) else None
    return out

# --------------------------- UI --------------------------------
st.set_page_config(page_title="Lector de Matrices → Resumen Excel", layout="wide")
st.title("📊 Lector de Matrices (Excel) → Resumen consolidado")

st.markdown("""
Sube tus matrices (.xlsx / .xlsm). La app detecta:
- **Delegación**, **Líneas de Acción**
- **Avance de Indicadores** (*Completos / Con actividades / Sin actividades*, con **n** y **%**)
- **Indicadores** por **Gobierno Local** y **Fuerza Pública**  
y genera un **Excel consolidado**.
""")

uploads = st.file_uploader("Arrastra o selecciona tus matrices", type=["xlsx","xlsm"], accept_multiple_files=True)

rows, failed = [], []
if uploads:
    for f in uploads:
        try:
            rows.append(process_file(f))
        except Exception as e:
            failed.append((f.name, str(e)))

    if rows:
        df_out = pd.DataFrame(rows)
        rename = {
            "archivo":"Archivo",
            "delegacion":"Delegación",
            "lineas_accion":"Líneas de Acción",
            "completos_n":"Completos (n)",
            "completos_pct":"Completos (%)",
            "conact_n":"Con actividades (n)",
            "conact_pct":"Con actividades (%)",
            "sinact_n":"Sin actividades (n)",
            "sinact_pct":"Sin actividades (%)",
            "indicadores_gl":"Indicadores Gobierno Local",
            "indicadores_fp":"Indicadores Fuerza Pública",
            "indicadores_total":"Total Indicadores",
        }
        df_out = df_out[list(rename.keys())].rename(columns=rename)

        for col in ["Completos (%)","Con actividades (%)","Sin actividades (%)"]:
            if col in df_out.columns:
                df_out[col] = df_out[col].apply(lambda v: f"{v:.1f}%" if pd.notna(v) else None)

        st.subheader("Resumen previo")
        st.dataframe(df_out, use_container_width=True)

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

