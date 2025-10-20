# -*- coding: utf-8 -*-
# ================================================================
# Lector de Matrices (Excel) → Resumen consolidado en Excel
# - Soporta múltiples .xlsx/.xlsm
# - Detección por rótulos (no por posiciones)
# - Filtros anti falsos positivos (años, % como cantidades, etc.)
# - Avance (n y %), GL/FP, Total de Indicadores
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
    # extrae sólo dígitos y signo y valida 1..3 dígitos
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
    # header=None mantiene todo como data; dtype=str conserva formatos (%, etc.)
    return pd.read_excel(file, engine="openpyxl", header=None, dtype=str)

def _find(df: pd.DataFrame, pattern: str) -> List[Tuple[int,int]]:
    """Encuentra celdas cuyo texto normalizado matchee el patrón (regex)."""
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
    """Filtra None y fuera de rango; prioriza el mayor típico (evita 1 residuales)."""
    cands = [x for x in cands if x is not None and 0 <= x <= max_allowed]
    if not cands: 
        return None
    return max(cands)

# ------------------- Detectores por rótulos ---------------------
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
    hits = _find(df, r"\blineas?\s*de\s*accion\b")
    for (r,c) in hits:
        cands = []
        for (i,j) in _neighbors(df, r, c, up=0, down=6, left=2, right=4):
            cands.append(_to_int(df.iat[i,j]))
        nums = [v for v in cands if v is not None]
        if nums:
            val = _pick_best_count(nums, max_allowed=60)
            if debug and val is not None:
                st.caption(f"Líneas de Acción detectadas: {val}")
            return val, None
    return None, None

def _row_has_all(row_vals: List[str], needed: List[str]) -> bool:
    normed = [_norm(v) for v in row_vals]
    return all(any(k in cell for cell in normed) for k in needed)

def detect_avance_indicadores(df: pd.DataFrame, debug: bool=False) -> Dict[str, Dict[str, Optional[float]]]:
    """
    Devuelve:
      {'completos': {'n','%'}, 'con_actividades': {'n','%'}, 'sin_actividades': {'n','%'}}
    - Encuentra fila con 'Completos | Con actividades | Sin actividades'
    - Toma la primera fila numérica (sin %) debajo y la primera fila con % debajo
    - Tolera mezcla n/% entre esas dos filas
    """
    res = { "completos":{"n":None,"%":None},
            "con_actividades":{"n":None,"%":None},
            "sin_actividades":{"n":None,"%":None} }

    for r in range(df.shape[0]):
        row = [str(x) if x is not None else "" for x in df.iloc[r,:].tolist()]
        if not row:
            continue
        normed = [_norm(x) for x in row]
        if not (any("complet" in x for x in normed) and
                any("con actividades" in x for x in normed) and
                any("sin actividades" in x for x in normed)):
            continue

        # Mapear columna para cada encabezado
        hdr_idx = {}
        for j, h in enumerate(normed):
            if "complet" in h and "completos" not in hdr_idx:
                hdr_idx["completos"] = j
            elif "con actividades" in h and "con_actividades" not in hdr_idx:
                hdr_idx["con_actividades"] = j
            elif "sin actividades" in h and "sin_actividades" not in hdr_idx:
                hdr_idx["sin_actividades"] = j

        # Primera fila numérica (sin %) debajo
        num_row = None
        rr = r + 1
        while rr < df.shape[0]:
            vals = [str(x) if x is not None else "" for x in df.iloc[rr,:].tolist()]
            has_pct = any("%" in v for v in vals)
            has_int = any(_to_int(v) is not None for v in vals)
            if has_int and not has_pct:
                num_row = vals
                break
            rr += 1

        # Primera fila de porcentajes (con %) debajo
        pct_row = None
        rr2 = r + 1
        while rr2 < df.shape[0]:
            vals = [str(x) if x is not None else "" for x in df.iloc[rr2,:].tolist()]
            if any("%" in v for v in vals):
                pct_row = vals
                break
            rr2 += 1

        for key, j in hdr_idx.items():
            n = _to_int(num_row[j]) if (num_row and j < len(num_row)) else None
            p = _to_pct(pct_row[j]) if (pct_row and j < len(pct_row)) else None
            # Mezcla tolerada
            if n is None and num_row and j < len(num_row):
                p2 = _to_pct(num_row[j])
                if p is None and p2 is not None:
                    p = p2
            if p is None and pct_row and j < len(pct_row):
                n2 = _to_int(pct_row[j])
                if n is None and n2 is not None:
                    n = n2
            res[key] = {"n": n, "%": p}

        if debug:
            st.caption(f"Avance detectado en fila {r}: {res}")
        return res

    return res

def detect_indicadores_categoria(df: pd.DataFrame, categoria: str, debug: bool=False) -> Tuple[Optional[int], Optional[str]]:
    """
    1) Busca el texto 'categoria' (Gobierno Local | Fuerza Pública)
    2) Localiza 'Indicadores' cerca y toma el número grande adyacente (arriba/abajo ±1 col)
    3) Si falla, suma la última cifra entera en filas con % (tablita de 3 filas)
    """
    anchors = _find(df, _norm(categoria))
    if not anchors:
        return None, "sin_ancla"

    for (ra, ca) in anchors:
        # (1) Buscar rótulo "Indicadores" cerca
        inds = []
        for (ri, ci) in _neighbors(df, ra, ca, up=0, down=18, left=0, right=18):
            if "indicadores" in _norm(df.iat[ri, ci]):
                inds.append((ri, ci))

        # (2) Número grande adyacente
        for (ri, ci) in inds:
            cands = []
            for delta_r in (-2, -1, 1, 2):
                rr = ri + delta_r
                if 0 <= rr < df.shape[0]:
                    for dc in (-1, 0, 1):
                        cc = ci + dc
                        if 0 <= cc < df.shape[1]:
                            cands.append(_to_int(df.iat[rr, cc]))
            val = _pick_best_count([v for v in cands if v is not None], max_allowed=60)
            if val is not None:
                if debug: st.caption(f"{categoria}: número grande cerca de 'Indicadores' → {val}")
                return val, "indicadores_label"

        # (3) Suma de última cifra entera por fila (en filas que tengan %)
        sum_cand, found = 0, False
        for rr in range(ra, min(ra + 15, df.shape[0])):
            row = [df.iat[rr, cc] for cc in range(max(0, ca - 5), min(ca + 10, df.shape[1]))]
            if not row: 
                continue
            if not any("%" in str(v) for v in row if v is not None):
                continue
            last_int = None
            for v in reversed(row):
                vi = _to_int(v)
                if vi is not None:
                    last_int = vi
                    break
            if last_int is not None:
                sum_cand += last_int
                found = True
        if found and 0 < sum_cand <= 60:
            if debug: st.caption(f"{categoria}: suma de tablita = {sum_cand}")
            return sum_cand, "tabla_sum"

    return None, "no_encontrado"

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

    lineas, _ = detect_lineas_accion(df, debug=debug)
    avance = detect_avance_indicadores(df, debug=debug)
    gl, _ = detect_indicadores_categoria(df, "gobierno local", debug=debug)
    fp, _ = detect_indicadores_categoria(df, "fuerza publica", debug=debug)
    total = detect_total_indicadores(df)

    # Cross-check básico con Total de Indicadores
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

        # Formato de porcentaje
        for col in ["Completos (%)","Con actividades (%)","Sin actividades (%)"]:
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

