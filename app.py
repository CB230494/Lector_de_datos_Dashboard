# -*- coding: utf-8 -*-
# ================================================================
# Lector de Matrices (Excel) → Resumen consolidado en Excel
# - Soporta múltiples archivos .xlsx / .xlsm
# - Detección por etiquetas (robusta) con modo de rescate manual
# - Extrae: Delegación, Líneas de Acción, Avance de Indicadores,
#           Indicadores de Gobierno Local y Fuerza Pública
# ================================================================

import io
import re
import unicodedata
from typing import Dict, Optional, Tuple, List

import numpy as np
import pandas as pd
import streamlit as st

# -------------------------------------------
# Utilidades de normalización y búsqueda
# -------------------------------------------
def _norm_text(x: str) -> str:
    if x is None:
        return ""
    x = str(x)
    x = unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", x).strip().lower()

def _coerce_int(x) -> Optional[int]:
    if x is None: 
        return None
    s = str(x)
    m = re.search(r"-?\d+", s.replace(",", ""))
    return int(m.group()) if m else None

def _coerce_pct(x) -> Optional[float]:
    """Devuelve porcentaje entre 0 y 100 si parece %; si trae 0-1 lo escala."""
    if x is None:
        return None
    s = str(x).replace(",", ".")
    m = re.search(r"-?\d+(\.\d+)?", s)
    if not m:
        return None
    val = float(m.group())
    if "%" in s or val > 1.0:
        return val  # ya está 0–100
    return val * 100.0  # venía 0–1

def _read_first_sheet(file) -> pd.DataFrame:
    # dtype=str para preservar textos y signos de %
    return pd.read_excel(file, engine="openpyxl", header=None, dtype=str)

def _find_cells(df: pd.DataFrame, pattern: str) -> List[Tuple[int, int]]:
    """Encuentra celdas cuyo texto normalizado haga match con pattern (regex en minúsculas)."""
    coords = []
    rx = re.compile(pattern)
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            txt = _norm_text(df.iat[r, c])
            if txt and rx.search(txt):
                coords.append((r, c))
    return coords

def _first_nonempty_numeric_below(df: pd.DataFrame, r: int, c: int, max_down: int = 5) -> Optional[int]:
    for k in range(1, max_down + 1):
        if r + k < df.shape[0]:
            val = _coerce_int(df.iat[r + k, c])
            if val is not None:
                return val
    # prueba también misma fila a la derecha
    for k in range(1, 4):
        if c + k < df.shape[1]:
            val = _coerce_int(df.iat[r, c + k])
            if val is not None:
                return val
    return None

# -------------------------------------------
# Detectores específicos por bloque
# -------------------------------------------
def detect_delegacion(df: pd.DataFrame) -> Optional[str]:
    # Busca patrones tipo D35-Orotina, D35 – Orotina, etc.
    for r in range(min(10, df.shape[0])):  # usualmente está arriba
        for c in range(df.shape[1]):
            raw = str(df.iat[r, c]) if df.iat[r, c] is not None else ""
            norm = _norm_text(raw)
            if re.match(r"^d\d{1,3}\s*[-–]\s*.+$", norm):
                # Devuelve el texto original sin normalizar
                return raw.strip()
    # Fallback: busca "delegacion" y toma celda a la derecha
    hits = _find_cells(df, r"\bdelegaci[oó]n\b")
    for (r, c) in hits:
        if c + 1 < df.shape[1]:
            raw = df.iat[r, c + 1]
            if raw and str(raw).strip():
                return str(raw).strip()
    return None

def detect_lineas_accion(df: pd.DataFrame) -> Optional[int]:
    hits = _find_cells(df, r"lineas?\s*de\s*accion")
    for (r, c) in hits:
        val = _first_nonempty_numeric_below(df, r, c, max_down=6)
        if val is not None:
            return val
    # Fallback: busca cualquier número grande cerca del título
    return None

def _row_has_headers(row_vals: List[str], headers: List[str]) -> bool:
    norm_row = [_norm_text(x) for x in row_vals]
    return all(any(h in cell for cell in norm_row) for h in headers)

def detect_avance_indicadores(df: pd.DataFrame) -> Dict[str, Dict[str, Optional[float]]]:
    """
    Devuelve:
    {
      'completos': {'n': int, '%': float},
      'con_actividades': {'n': int, '%': float},
      'sin_actividades': {'n': int, '%': float}
    }
    """
    res = {
        "completos": {"n": None, "%": None},
        "con_actividades": {"n": None, "%": None},
        "sin_actividades": {"n": None, "%": None},
    }

    # 1) ubica bloque por cabecera
    for r in range(df.shape[0]):
        row_vals = [str(x) if x is not None else "" for x in df.iloc[r, :].tolist()]
        if _row_has_headers(row_vals, ["avance de indicadores"]):
            # busca una fila siguiente que contenga los tres encabezados
            for rr in range(r, min(r + 10, df.shape[0])):
                row_vals2 = [str(x) if x is not None else "" for x in df.iloc[rr, :].tolist()]
                if _row_has_headers(row_vals2, ["complet", "con actividades", "sin actividades"]):
                    # Se asume: rr es headers; rr+1 = números, rr+2 = porcentajes (o mezclado)
                    nums = [str(x) if x is not None else "" for x in df.iloc[rr + 1, :].tolist()] if rr + 1 < df.shape[0] else []
                    pcts = [str(x) if x is not None else "" for x in df.iloc[rr + 2, :].tolist()] if rr + 2 < df.shape[0] else []

                    def pick(col_name: str) -> Tuple[Optional[int], Optional[float]]:
                        # encuentra índice por cabecera aproximada
                        hdrs = [str(x) if x is not None else "" for x in df.iloc[rr, :].tolist()]
                        idx = None
                        for j, h in enumerate(hdrs):
                            if col_name in _norm_text(h):
                                idx = j
                                break
                        if idx is None:
                            return None, None
                        n = _coerce_int(nums[idx]) if idx < len(nums) else None
                        p = _coerce_pct(pcts[idx]) if idx < len(pcts) else None

                        # Si la “fila de números” traía % mezclado, corrige
                        if (n is None) and idx < len(nums):
                            maybe_pct = _coerce_pct(nums[idx])
                            if maybe_pct is not None:
                                p = maybe_pct
                        if (p is None) and idx < len(pcts):
                            maybe_int = _coerce_int(pcts[idx])
                            if maybe_int is not None:
                                n = maybe_int
                        return n, p

                    n, p = pick("complet")
                    res["completos"] = {"n": n, "%": p}
                    n, p = pick("con actividades")
                    res["con_actividades"] = {"n": n, "%": p}
                    n, p = pick("sin actividades")
                    res["sin_actividades"] = {"n": n, "%": p}
                    return res
    # Fallback: escaneo libre por fila con esos encabezados (sin "avance de indicadores")
    for r in range(df.shape[0]):
        row_vals = [str(x) if x is not None else "" for x in df.iloc[r, :].tolist()]
        if _row_has_headers(row_vals, ["complet", "con actividades", "sin actividades"]):
            nums = [str(x) if x is not None else "" for x in df.iloc[r + 1, :].tolist()] if r + 1 < df.shape[0] else []
            pcts = [str(x) if x is not None else "" for x in df.iloc[r + 2, :].tolist()] if r + 2 < df.shape[0] else []
            def pick_idx(target):
                hdrs = [str(x) if x is not None else "" for x in df.iloc[r, :].tolist()]
                for j, h in enumerate(hdrs):
                    if target in _norm_text(h):
                        return j
                return None
            for key, target in [("completos","complet"), ("con_actividades","con actividades"), ("sin_actividades","sin actividades")]:
                j = pick_idx(target)
                if j is not None:
                    n = _coerce_int(nums[j]) if j < len(nums) else None
                    p = _coerce_pct(pcts[j]) if j < len(pcts) else None
                    if (n is None) and j < len(nums):
                        maybe_pct = _coerce_pct(nums[j])
                        if maybe_pct is not None:
                            p = maybe_pct
                    if (p is None) and j < len(pcts):
                        maybe_int = _coerce_int(pcts[j])
                        if maybe_int is not None:
                            n = maybe_int
                    res[key] = {"n": n, "%": p}
            return res
    return res

def detect_indicadores_categoria(df: pd.DataFrame, titulo: str) -> Optional[int]:
    """
    Dado 'Gobierno Local' o 'Fuerza Pública', busca el pequeño cuadro de % y cantidad.
    Estrategia: suma de las cantidades de la tablita, o busca un número grande
    cercano al texto 'Indicadores'.
    """
    hits = _find_cells(df, _norm_text(titulo))
    for (r, c) in hits:
        # 1) intenta sumar cantidades en las próximas 5 filas y 3 columnas (formato usual)
        cant_sum = 0
        found_any = False
        for rr in range(r + 1, min(r + 6, df.shape[0])):
            # suele ser [porcentaje | cantidad]
            row_vals = [df.iat[rr, cc] if cc < df.shape[1] else None for cc in range(c, min(c + 4, df.shape[1]))]
            # toma el último número entero de la fila como "cantidad"
            nums = [_coerce_int(v) for v in row_vals]
            nums = [n for n in nums if n is not None]
            if nums:
                cant_sum += nums[-1]
                found_any = True
        if found_any and cant_sum > 0:
            return cant_sum

        # 2) busca la palabra "indicadores" debajo y toma número cercano
        for rr in range(r, min(r + 10, df.shape[0])):
            for cc in range(max(0, c - 5), min(df.shape[1], c + 6)):
                if "indicadores" in _norm_text(df.iat[rr, cc]):
                    # número a la izquierda o derecha inmediata
                    cand = []
                    if cc - 1 >= 0:
                        cand.append(_coerce_int(df.iat[rr, cc - 1]))
                    cand.append(_coerce_int(df.iat[rr, cc + 1] if cc + 1 < df.shape[1] else None))
                    cand = [x for x in cand if x is not None]
                    if cand:
                        return max(cand)
    return None

# -------------------------------------------
# Procesar un archivo
# -------------------------------------------
def process_file(upload) -> Dict:
    df = _read_first_sheet(upload)
    out = {
        "archivo": upload.name,
        "delegacion": detect_delegacion(df),
        "lineas_accion": detect_lineas_accion(df),
        "completos_n": None, "completos_pct": None,
        "conact_n": None, "conact_pct": None,
        "sinact_n": None, "sinact_pct": None,
        "indicadores_gl": detect_indicadores_categoria(df, "Gobierno Local"),
        "indicadores_fp": detect_indicadores_categoria(df, "Fuerza Publica"),
    }
    av = detect_avance_indicadores(df)
    out["completos_n"] = av["completos"]["n"]
    out["completos_pct"] = av["completos"]["%"]
    out["conact_n"] = av["con_actividades"]["n"]
    out["conact_pct"] = av["con_actividades"]["%"]
    out["sinact_n"] = av["sin_actividades"]["n"]
    out["sinact_pct"] = av["sin_actividades"]["%"]
    # Totales
    gl = out["indicadores_gl"] or 0
    fp = out["indicadores_fp"] or 0
    out["indicadores_total"] = gl + fp if (gl or fp) else None
    return out

# -------------------------------------------
# Streamlit UI
# -------------------------------------------
st.set_page_config(page_title="Lector de Matrices → Resumen Excel", layout="wide")
st.title("📊 Lector de Matrices (Excel) → Resumen consolidado")

st.markdown("""
Sube tus matrices (.xlsx / .xlsm). La app detecta:
- **Delegación**, **Líneas de Acción**
- **Avance de Indicadores** (*Completos / Con actividades / Sin actividades*, con **n** y **%**)
- **Indicadores** por **Gobierno Local** y **Fuerza Pública**  
Luego podrás **descargar** un Excel con el consolidado.
""")

with st.sidebar:
    st.header("Opciones")
    st.caption("Si algún archivo no sigue la plantilla exacta, puedes usar el modo de rescate para ingresar celdas manuales.")
    rescue = st.toggle("Activar modo de rescate manual (por archivo)", value=False)

uploads = st.file_uploader("Arrastra o selecciona los archivos de tus matrices", type=["xlsx", "xlsm"], accept_multiple_files=True)

rows: List[Dict] = []
failed_files: List[str] = []

if uploads:
    for f in uploads:
        try:
            info = process_file(f)
            rows.append(info)
        except Exception as e:
            failed_files.append(f.name)
            st.warning(f"⚠️ No se pudo procesar automáticamente: {f.name}. Error: {e}")

    # Vista previa
    if rows:
        df_out = pd.DataFrame(rows)
        # Reordenar y renombrar columnas a las solicitadas
        rename_map = {
            "archivo": "Archivo",
            "delegacion": "Delegación",
            "lineas_accion": "Líneas de Acción",
            "completos_n": "Completos (n)",
            "completos_pct": "Completos (%)",
            "conact_n": "Con actividades (n)",
            "conact_pct": "Con actividades (%)",
            "sinact_n": "Sin actividades (n)",
            "sinact_pct": "Sin actividades (%)",
            "indicadores_gl": "Indicadores Gobierno Local",
            "indicadores_fp": "Indicadores Fuerza Pública",
            "indicadores_total": "Total Indicadores",
        }
        cols = list(rename_map.keys())
        df_out = df_out[cols].rename(columns=rename_map)

        # Formato porcentajes a 0.0%
        for col in ["Completos (%)", "Con actividades (%)", "Sin actividades (%)"]:
            if col in df_out.columns:
                df_out[col] = df_out[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else None)

        st.subheader("Resumen previo (esto será lo que descargues)")
        st.dataframe(df_out, use_container_width=True)

        # Descargar Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, sheet_name="resumen")
        st.download_button(
            label="⬇️ Descargar Excel consolidado",
            data=buffer.getvalue(),
            file_name="resumen_matrices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Archivos con fallo automático → modo rescate
    if rescue and failed_files:
        st.divider()
        st.subheader("🛟 Modo de rescate manual")
        st.caption("Para cada archivo fallido, especifica las celdas (por ejemplo, B4, C12, etc.). Si no tienes algún dato, déjalo vacío.")

        for f in failed_files:
            with st.expander(f"Corregir: {f}"):
                col1, col2, col3 = st.columns(3)
                with col1:
                    cel_deleg = st.text_input(f"[{f}] Celda Delegación", key=f"{f}-deleg")
                    cel_lineas = st.text_input(f"[{f}] Celda Líneas de Acción", key=f"{f}-lineas")
                with col2:
                    cel_comp_n = st.text_input(f"[{f}] Celda Completos (n)", key=f"{f}-compn")
                    cel_comp_p = st.text_input(f"[{f}] Celda Completos (%)", key=f"{f}-compp")
                    cel_con_n = st.text_input(f"[{f}] Celda Con actividades (n)", key=f"{f}-conn")
                    cel_con_p = st.text_input(f"[{f}] Celda Con actividades (%)", key=f"{f}-conp")
                with col3:
                    cel_sin_n = st.text_input(f"[{f}] Celda Sin actividades (n)", key=f"{f}-sinn")
                    cel_sin_p = st.text_input(f"[{f}] Celda Sin actividades (%)", key=f"{f}-sinp")
                    cel_gl = st.text_input(f"[{f}] Celda Indicadores Gobierno Local (total)", key=f"{f}-gl")
                    cel_fp = st.text_input(f"[{f}] Celda Indicadores Fuerza Pública (total)", key=f"{f}-fp")

                sub = st.file_uploader(f"Vuelve a cargar el archivo {f} para leer esas celdas", type=["xlsx", "xlsm"], key=f"{f}-reup")
                if sub is not None and st.button(f"Tomar valores de celdas para {f}", key=f"{f}-btn"):
                    try:
                        df = _read_first_sheet(sub)
                        def at(cell):
                            m = re.match(r"^\s*([A-Za-z]+)(\d+)\s*$", cell or "")
                            if not m:
                                return None
                            col_letters, row_num = m.groups()
                            # convertir letras a índice 0-based
                            col_idx = 0
                            for ch in col_letters.upper():
                                col_idx = col_idx * 26 + (ord(ch) - ord('A') + 1)
                            col_idx -= 1
                            row_idx = int(row_num) - 1
                            if 0 <= row_idx < df.shape[0] and 0 <= col_idx < df.shape[1]:
                                return df.iat[row_idx, col_idx]
                            return None

                        manual = {
                            "archivo": f,
                            "Delegación": (at(cel_deleg) or "").strip() if cel_deleg else None,
                            "Líneas de Acción": _coerce_int(at(cel_lineas)) if cel_lineas else None,
                            "Completos (n)": _coerce_int(at(cel_comp_n)) if cel_comp_n else None,
                            "Completos (%)": _coerce_pct(at(cel_comp_p)) if cel_comp_p else None,
                            "Con actividades (n)": _coerce_int(at(cel_con_n)) if cel_con_n else None,
                            "Con actividades (%)": _coerce_pct(at(cel_con_p)) if cel_con_p else None,
                            "Sin actividades (n)": _coerce_int(at(cel_sin_n)) if cel_sin_n else None,
                            "Sin actividades (%)": _coerce_pct(at(cel_sin_p)) if cel_sin_p else None,
                            "Indicadores Gobierno Local": _coerce_int(at(cel_gl)) if cel_gl else None,
                            "Indicadores Fuerza Pública": _coerce_int(at(cel_fp)) if cel_fp else None,
                        }
                        gl = manual["Indicadores Gobierno Local"] or 0
                        fp = manual["Indicadores Fuerza Pública"] or 0
                        manual["Total Indicadores"] = gl + fp if (gl or fp) else None

                        st.success("Valores tomados. Agrégalos manualmente al consolidado descargado o vuelve a procesar estos archivos con las mismas celdas.")
                        st.json(manual)
                    except Exception as e:
                        st.error(f"No se pudo leer el archivo/celdas: {e}")

else:
    st.info("Sube tus matrices para ver el resumen.")

