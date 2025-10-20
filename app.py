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
    s = str(x).strip()
    if "%" in s:              # nunca usamos % como cantidades
        return None
    m = re.fullmatch(r"-?\d{1,3}", re.sub(r"[^\d-]", "", s))  # 1 a 3 dígitos
    if m:
        try:
            val = int(m.group())
            return val
        except:
            return None
    return None  # descarta 2025 y similares

def _to_pct(x) -> Optional[float]:
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
    # preferimos el más grande típico (p. ej. 4 frente a 1 residuales)
    return max(cands)

# ------------------- Detectores robustos ------------------------
def detect_delegacion(df: pd.DataFrame) -> Optional[str]:
    # Ej: "D35-Orotina"
    rx = re.compile(r"^\s*d\d{1,3}\s*[-–]\s*.+\s*$", re.IGNORECASE)
    for r in range(min(15, df.shape[0])):
        for c in range(df.shape[1]):
            raw = df.iat[r,c]
            if raw and rx.match(str(raw)):
                return str(raw).strip()
    # Fallback por rótulo "delegación"
    hits = _find(df, r"\bdelegaci[oó]n\b")
    for (r,c) in hits:
        if c+1 < df.shape[1]:
            v = df.iat[r, c+1]
            if v: return str(v).strip()
    return None

def detect_lineas_accion(df: pd.DataFrame, debug: bool=False) -> Tuple[Optional[int], Optional[Tuple[int,int]]]:
    # Ancla por rótulo
    hits = _find(df, r"\blineas?\s*de\s*accion\b")
    for (r,c) in hits:
        # Busca cerca del rótulo números candidatos 0..60 sin %
        cands, pos = [], None
        for (i,j) in _neighbors(df, r, c, up=0, down=6, left=2, right=4):
            val = _to_int(df.iat[i,j])
            if val is not None:
                cands.append((val, (i,j)))
        if cands:
            val, pos = max(cands, key=lambda t: t[0])
            if debug: st.caption(f"Líneas de Acción tomado en {pos} = {val}")
            return val, pos
    return None, None

def _row_has_all(row_vals: List[str], needed: List[str]) -> bool:
    normed = [_norm(v) for v in row_vals]
    return all(any(k in cell for cell in normed) for k in needed)

def detect_avance_indicadores(df: pd.DataFrame, debug: bool=False) -> Dict[str, Dict[str, Optional[float]]]:
    res = { "completos":{"n":None,"%":None},
            "con_actividades":{"n":None,"%":None},
            "sin_actividades":{"n":None,"%":None} }
    for r in range(df.shape[0]):
        hdr = [str(x) if x is not None else "" for x in df.iloc[r,:].tolist()]
        if _row_has_all(hdr, ["complet", "con actividades", "sin actividades"]):
            nums = [str(x) if x is not None else "" for x in df.iloc[r+1,:].tolist()] if r+1<df.shape[0] else []
            pcts = [str(x) if x is not None else "" for x in df.iloc[r+2,:].tolist()] if r+2<df.shape[0] else []
            def pick(key):
                idx = None
                for j,h in enumerate(hdr):
                    if key in _norm(h):
                        idx = j; break
                if idx is None: return None, None
                n = _to_int(nums[idx]) if idx < len(nums) else None
                p = _to_pct(pcts[idx]) if idx < len(pcts) else None
                if n is None and idx < len(nums):
                    p2 = _to_pct(nums[idx])
                    if p2 is not None: p = p2
                if p is None and idx < len(pcts):
                    n2 = _to_int(pcts[idx])
                    if n2 is not None: n = n2
                return n, p
            res["completos"]["n"], res["completos"]["%"] = pick("complet")
            res["con_actividades"]["n"], res["con_actividades"]["%"] = pick("con actividades")
            res["sin_actividades"]["n"], res["sin_actividades"]["%"] = pick("sin actividades")
            if debug: st.caption(f"Avance fila {r}: {res}")
            return res
    return res

def detect_indicadores_categoria(df: pd.DataFrame, categoria: str, debug: bool=False) -> Tuple[Optional[int], Optional[str]]:
    """
    1) Busca el texto 'categoria' (Gobierno Local | Fuerza Pública)
    2) Dentro del vecindario, localiza el rótulo 'Indicadores' y toma el número **encima** (misma o ±1 col)
       como candidato (0..60, entero sin %).
    3) Si no logra, intenta sumar la **última columna** de la tablita (que contiene cantidades puras).
    """
    anchors = _find(df, _norm(categoria))
    if not anchors:
        return None, "sin_ancla"

    best_val, best_note = None, None
    for (ra, ca) in anchors:
        # Paso 2: “Indicadores” cercano
        inds = []
        for (ri, ci) in _neighbors(df, ra, ca, up=0, down=15, left=0, right=12):
            if "indicadores" in _norm(df.iat[ri, ci]):
                inds.append((ri, ci))
        for (ri, ci) in inds:
            cands = []
            for up in (1,2,3):
                rr = ri - up
                if rr < 0: break
                for dc in (-1,0,1):
                    cc = ci + dc
                    if 0 <= cc < df.shape[1]:
                        cands.append(_to_int(df.iat[rr, cc]))
            val = _pick_best_count(cands, max_allowed=60)
            if val is not None:
                if debug: st.caption(f"{categoria}: Indicadores en {(ri,ci)} → {val}")
                return val, "indicadores_label"

        # Paso 3: sumar última col de la tablita (si existe)
        # Buscamos filas con patrones de porcentaje en primera/segunda col y número puro al final
        sum_cand = 0
        found = False
        for rr in range(ra, min(ra+12, df.shape[0])):
            row = [df.iat[rr, cc] for cc in range(ca, min(ca+6, df.shape[1]))]
            # detecta si hay un % en la fila (propio de esa tablita) y un entero al final
            has_pct = any("%" in str(v) for v in row if v is not None)
            last_int = None
            for v in reversed(row):
                last_int = _to_int(v)
                if last_int is not None:
                    break
            if has_pct and last_int is not None:
                found = True
                sum_cand += last_int
        if found and 0 < sum_cand <= 60:
            if debug: st.caption(f"{categoria}: suma de tablita = {sum_cand}")
            return sum_cand, "tabla_sum"

    return best_val, best_note

def detect_total_indicadores(df: pd.DataFrame) -> Optional[int]:
    hits = _find(df, r"\btotal\s+de\s+indicadores\b")
    for (r,c) in hits:
        cands = []
        for (i,j) in _neighbors(df, r, c, up=0, down=2, left=0, right=4):
            cands.append(_to_int(df.iat[i,j]))
        val = _pick_best_count(cands, max_allowed=120)
        if val is not None:
            return val
    return None

# --------------------- Proceso de un archivo --------------------
def process_file(upload, debug: bool=False) -> Dict:
    df = _read_df(upload)

    lineas, _ = detect_lineas_accion(df, debug=debug)
    avance = detect_avance_indicadores(df, debug=debug)
    gl, gl_note = detect_indicadores_categoria(df, "gobierno local", debug=debug)
    fp, fp_note = detect_indicadores_categoria(df, "fuerza publica", debug=debug)
    total = detect_total_indicadores(df)

    # Cross-check simple
    if total is not None and (gl is None or fp is None):
        if gl is not None and fp is None:
            fp = total - gl
        elif fp is not None and gl is None:
            gl = total - fp

    out = {
        "archivo": upload.name,
        "delegacion": detect_delegacion(df),
        "lineas_accion": lineas,

        "completos_n": avance["completos"]["n"],
        "completos_pct": avance["completos"]["%"],
        "conact_n": avance["con_actividades"]["n"],
        "conact_pct": avance["con_actividades"]["%"],
        "sinact_n": avance["sin_actividades"]["n"],
        "sinact_pct": avance["sin_actividades"]["%"],

        "indicadores_gl": gl,
        "indicadores_fp": fp,
        "indicadores_total": (gl or 0) + (fp or 0) if (gl is not None or fp is not None) else total,
    }
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
- **Indicadores** por **Gobierno Local** y **Fuerza Pública**  
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
        order = list(rename.keys())
        df_out = df_out[order].rename(columns=rename)

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
