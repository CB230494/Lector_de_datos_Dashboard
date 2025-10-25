# =========================
# PARTE 1/10 — BASE & UTILS
# =========================
# - Imports
# - Bandera de features
# - Normalización de texto y parsers seguros
# - Utilidades de búsqueda en celdas
# - Helpers genéricos (tokens, números) reutilizados por todo el módulo

# -*- coding: utf-8 -*-
import io, re, unicodedata
from typing import Dict, Optional, Tuple, List
import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# -------------------------------------------------------------------
# Configuración de visualización (banderas que controlan features)
# -------------------------------------------------------------------
# Si lo pones True, se mostrará el panel extra de filtros (categorías/problema/responsable).
SHOW_EXTRA_FILTERS = False

# -------------------------------------------------------------------
# Normalización / parsing robusto
# -------------------------------------------------------------------
def _norm(x: str) -> str:
    """Normaliza texto a ASCII, colapsa espacios y pasa a minúsculas."""
    if x is None:
        return ""
    x = str(x)
    x = unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", x).strip().lower()

def _norm_str(s: str) -> str:
    """Versión 'suave' para nombres de columnas/etiquetas (conserva palabras)."""
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    s = re.sub(r"[\s_\-]+", " ", s).strip().lower()
    return s

def _to_int(x) -> Optional[int]:
    """
    Convierte a entero solo si parece una cantidad corta (1–3 dígitos).
    Evita años (p.ej., 2025) y celdas con '%'.
    """
    if x is None:
        return None
    s = str(x).strip()
    if "%" in s:
        return None
    digits = re.sub(r"[^\d-]", "", s)
    if not re.fullmatch(r"-?\d{1,3}", digits):
        return None
    try:
        return int(digits)
    except Exception:
        return None

def _to_pct(x) -> Optional[float]:
    """
    Interpreta porcentajes robusto:
    - '45', '45%', 0.45 → 45.0
    - Clampa entre 0..100
    """
    if x is None:
        return None
    s = str(x).replace(",", ".")
    m = re.search(r"-?\d+(\.\d+)?", s)
    if not m:
        return None
    v = float(m.group())
    if "%" in s or v > 1.0:
        v = max(0.0, min(100.0, v))
    else:
        v = max(0.0, min(100.0, v * 100.0))
    return v

def _to_num_safe(x, pct: bool=False) -> float:
    """
    Conversión resiliente a número float.
    - pct=True: quita '%' y devuelve float 0..100 si es posible.
    - Si no se puede, intenta extraer primer número en el string.
    """
    if pd.isna(x):
        return 0.0
    s = str(x).strip()
    if pct:
        s = s.replace("%", "").replace(",", ".")
        try:
            return float(s)
        except Exception:
            return 0.0
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        m = re.search(r"-?\d+(\.\d+)?", s)
        return float(m.group()) if m else 0.0

# -------------------------------------------------------------------
# Lectura y búsqueda en matrices Excel
# -------------------------------------------------------------------
def _read_df(file) -> pd.DataFrame:
    """Lee Excel sin encabezados, todo como texto (para layouts variables)."""
    return pd.read_excel(file, engine="openpyxl", header=None, dtype=str)

def _find(df: pd.DataFrame, pattern: str) -> List[Tuple[int, int]]:
    """Encuentra celdas cuyo contenido (normalizado) case con el patrón regex."""
    rx = re.compile(pattern)
    out: List[Tuple[int, int]] = []
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = df.iat[r, c]
            if val is None:
                continue
            if rx.search(_norm(val)):
                out.append((r, c))
    return out

def _neighbors(df: pd.DataFrame, r: int, c: int, up: int, down: int, left: int, right: int):
    """Genera coordenadas (i, j) en la vecindad de (r, c) con límites seguros."""
    r0 = max(0, r - up)
    r1 = min(df.shape[0] - 1, r + down)
    c0 = max(0, c - left)
    c1 = min(df.shape[1] - 1, c + right)
    for i in range(r0, r1 + 1):
        for j in range(c0, c1 + 1):
            yield i, j

def _pick_best_count(cands: List[int], max_allowed: int = 60) -> Optional[int]:
    """Elige el mejor conteo plausible (0..max_allowed) de una lista de candidatos."""
    cands = [x for x in cands if x is not None and 0 <= x <= max_allowed]
    return max(cands) if cands else None

# -------------------------------------------------------------------
# Helpers para columnas "dinámicas" (Categorias / Problemática / etc.)
# -------------------------------------------------------------------
def _pick_named_column(df: pd.DataFrame, names_like: List[str]) -> Optional[str]:
    """
    Dada una lista de alias (p.ej., ['categorias','categoría','categoria']),
    intenta mapear a la columna real del df, robusto a tildes/espacios.
    """
    if df is None or df.empty:
        return None
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
    """
    Divide celdas tipo 'a, b / c' en tokens limpios.
    Separadores: coma, punto-y-coma, slash, pipe y saltos de línea.
    """
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
# ===========================
# PARTE 2/10 — DETECTORES
# ===========================
# - Delegación (por rótulo y por patrón "D## - Nombre")
# - Líneas de Acción (por rótulo + conteo cercano o por patrón "Línea de acción #")
# - Detección de filas de indicadores (GL/FP)
# - Detección de columnas de "Avance"
# - Conteos globales y por rol de avance (completos / con actividades / sin actividades)
# - Intento de leer "Total de indicadores" desde la matriz

def detect_delegacion(df: pd.DataFrame) -> Optional[str]:
    """
    Heurística de delegación:
    1) Busca patrón 'D{1..3} - ...' en primeras ~15 filas (cabeceras habituales).
    2) Si no, busca rótulo 'Delegación' y toma la celda a su derecha.
    """
    rx = re.compile(r"^\s*d\d{1,3}\s*[-–]\s*.+\s*$", re.IGNORECASE)
    for r in range(min(15, df.shape[0])):
        for c in range(df.shape[1]):
            raw = df.iat[r, c]
            if raw and rx.match(str(raw)):
                return str(raw).strip()
    hits = _find(df, r"\bdelegaci[oó]n\b")
    for (r, c) in hits:
        if c + 1 < df.shape[1]:
            v = df.iat[r, c + 1]
            if v:
                return str(v).strip()
    return None


def detect_lineas_accion(df: pd.DataFrame, debug: bool = False) -> Optional[int]:
    """
    Detecta cantidad de 'Líneas de Acción'.
    - Busca rótulo 'líneas de acción' y lee números cercanos (0..60).
    - Fallback: cuenta ocurrencias de 'Línea de acción #' si existen.
    """
    hits = _find(df, r"\blineas?\s*de\s*accion\b")
    for (r, c) in hits:
        cands: List[Optional[int]] = []
        for (i, j) in _neighbors(df, r, c, up=0, down=6, left=2, right=4):
            cands.append(_to_int(df.iat[i, j]))
        val = _pick_best_count(cands, max_allowed=60)
        if debug and val is not None:
            st.caption(f"Líneas de Acción detectadas: {val}")
        if val is not None:
            return val

    la_hits = _find(df, r"\blinea\s+de\s+accion\s*#")
    if la_hits:
        return len(la_hits)
    return None


# ------------------ GL/FP + AVANCE (por contenido visible) ------------------

def detect_indicator_rows(df: pd.DataFrame) -> List[int]:
    """
    Devuelve índices de filas que contienen 'GL' o 'FP' en las primeras ~6 columnas.
    Se asume que cada fila (indicador) está marcada con el rol en esa zona izquierda.
    """
    rows: List[int] = []
    for r in range(df.shape[0]):
        left_vals = [df.iat[r, c] for c in range(min(6, df.shape[1]))]
        left_norm = [_norm(v) for v in left_vals]
        if any(v == "gl" for v in left_norm) or any(v == "fp" for v in left_norm):
            rows.append(r)
    return rows


def detect_role_of_row(df: pd.DataFrame, r: int) -> Optional[str]:
    """
    Inspecciona las ~6 primeras columnas de la fila 'r'.
    Retorna: 'gl', 'fp' o None.
    """
    left_vals = [df.iat[r, c] for c in range(min(6, df.shape[1]))]
    left_norm = [_norm(v) for v in left_vals]
    if any(v == "gl" for v in left_norm):
        return "gl"
    if any(v == "fp" for v in left_norm):
        return "fp"
    return None


def detect_avance_columns(df: pd.DataFrame) -> List[int]:
    """
    Encuentra columnas cuyo encabezado/texto contiene 'Avance' (en filas superiores).
    Devuelve lista de índices. Si no detecta, usa fallback [10,15,20,25] válidos dentro del ancho.
    """
    cols: List[int] = []
    for r in range(df.shape[0]):
        row = [df.iat[r, c] for c in range(df.shape[1])]
        for c, v in enumerate(row):
            if "avance" in _norm(v):
                cols.append(c)
        if len(cols) >= 2:  # con 2 rotulados alcanzamos
            return sorted(list(set(cols)))
        else:
            cols = []
    fallback = [10, 15, 20, 25]
    return [c for c in fallback if c < df.shape[1]]


def gl_fp_counts(df: pd.DataFrame) -> Tuple[int, int]:
    """
    Cuenta filas GL y FP según 'detect_indicator_rows' + 'detect_role_of_row'.
    """
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
    """
    Dado el conjunto de columnas de avance, identifica el estado de la fila r:
    'completos' / 'con_actividades' / 'sin_actividades'
    (por búsqueda de palabras clave en esas columnas).
    """
    vals = [df.iat[r, c] for c in avance_cols]
    valsn = [_norm(v) for v in vals]
    if any("complet" in v for v in valsn):
        return "completos"
    if any("con actividades" in v for v in valsn):
        return "con_actividades"
    if any("sin actividades" in v for v in valsn):
        return "sin_actividades"
    # por defecto, conservador:
    return "sin_actividades"


def avance_counts(df: pd.DataFrame) -> Tuple[Dict[str, int], int]:
    """
    Conteo global de avance por estado para todas las filas GL/FP detectadas.
    Retorna (dict_estado, total_filas_detectadas).
    """
    rows = detect_indicator_rows(df)
    avance_cols = detect_avance_columns(df)
    counts = {"completos": 0, "con_actividades": 0, "sin_actividades": 0}
    for r in rows:
        counts[_row_status_from_avance(df, r, avance_cols)] += 1
    return counts, len(rows)


def avance_counts_by_role(df: pd.DataFrame) -> Tuple[Dict[str, int], int, Dict[str, int], int]:
    """
    Conteo de avance por estado, separado por rol GL y FP.
    Retorna (counts_gl, n_gl, counts_fp, n_fp).
    """
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
    """
    Intenta leer un rótulo 'Total de indicadores' y un número cercano (0..120).
    Útil cuando el archivo lo declara explícitamente.
    """
    hits = _find(df, r"\btotal\s+de\s+indicadores\b")
    for (r, c) in hits:
        cands: List[Optional[int]] = []
        for (i, j) in _neighbors(df, r, c, up=0, down=3, left=0, right=6):
            cands.append(_to_int(df.iat[i, j]))
        val = _pick_best_count([x for x in cands if x is not None], max_allowed=120)
        if val is not None:
            return val
    return None
# ================================
# PARTE 3/10 — PIPELINE POR ARCHIVO
# ================================
# - Lee el Excel (con _read_df de la Parte 1)
# - Detecta Delegación y Líneas de Acción
# - Calcula conteos globales y por rol (GL/FP)
# - Intenta inferir Total de Indicadores
# - Devuelve un dict (una fila) para el consolidado

def process_file(upload, debug: bool = False) -> Dict:
    df = _read_df(upload)

    # 1) Metadatos principales
    deleg = detect_delegacion(df)
    lineas = detect_lineas_accion(df, debug=debug)

    # 2) Conteos de filas por rol
    gl, fp = gl_fp_counts(df)

    # 3) Avance global (sobre todas las filas GL/FP)
    avance_dict, total_ind = avance_counts(df)
    comp_n = avance_dict["completos"]
    con_n  = avance_dict["con_actividades"]
    sin_n  = avance_dict["sin_actividades"]

    def pct(n, d):
        return round((n / d) * 100.0, 1) if d and n is not None else None

    comp_p = pct(comp_n, total_ind)
    con_p  = pct(con_n,  total_ind)
    sin_p  = pct(sin_n,  total_ind)

    # 4) Avance por rol
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

    # 5) Intento de total declarado en el archivo
    total_from_label = detect_total_indicadores(df)
    total_out = total_ind if total_ind else total_from_label

    # 6) Ajuste de gl/fp si falta un lado pero hay total (heurística conservadora)
    if (gl == 0 or fp == 0) and total_out and gl + fp != total_out:
        if gl == 0 and fp > 0:
            gl = max(0, total_out - fp)
        elif fp == 0 and gl > 0:
            fp = max(0, total_out - gl)

    # 7) Salida estándar para consolidado
    out = {
        "archivo": upload.name,
        "delegacion": deleg,
        "lineas_accion": lineas,

        # Avance global
        "completos_n": comp_n,
        "completos_pct": comp_p,
        "conact_n": con_n,
        "conact_pct": con_p,
        "sinact_n": sin_n,
        "sinact_pct": sin_p,

        # Indicadores por rol (conteo de filas)
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

        # Total inferido
        "indicadores_total": total_out if total_out is not None else (gl + fp if (gl or fp) else None),
    }

    if debug:
        st.caption(
            f"[DEBUG] Avance cols: {detect_avance_columns(df)} | "
            f"rows GL/FP detectadas: gl={gl} fp={fp} (total={gl+fp}) | "
            f"n_gl={n_gl} n_fp={n_fp}"
        )
    return out
# ============================================
# PARTE 4/10 — UI: Carga, consolidado y descarga
# ============================================

# Config general de la página
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

uploads = st.file_uploader(
    "Arrastra o selecciona tus matrices",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True
)

rows, failed = [], []
if uploads:
    for f in uploads:
        try:
            rows.append(process_file(f, debug=debug))
        except Exception as e:
            failed.append((f.name, str(e)))

    if rows:
        df_out = pd.DataFrame(rows)

        # Renombrado de columnas para el Excel final
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

        # Formatear % como texto con símbolo
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

        # Descargar Excel consolidado (hoja 'resumen')
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
# ==========================================
# PARTE 5/10 — HELPERS DE ESTILO Y NÚMEROS
# ==========================================

# Paleta de colores consistente
COLOR_ROJO   = "#ED1C24"
COLOR_AMARIL = "#F4C542"
COLOR_VERDE  = "#7AC943"
COLOR_AZUL_H = "#1F4E79"

def _big_number(value, label, helptext=None, big_px=80):
    """Componente simple de número grande con etiqueta."""
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

def _bar_avance(pcts_tuple, title=""):
    """Barra 3 estados: Sin actividades / Con actividades / Cumplida."""
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
    """Panel con tres tarjetas (rojo/amarillo/verde) y total abajo."""
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
    """Bloque 'Avance de Indicadores' con las tres tarjetas y el total."""
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

def _ensure_numeric(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte a numérico las columnas clave (cuentas y %).
    Usa _to_num_safe para tolerar strings y símbolos.
    """
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
# ==============================================
# PARTE 6/10 — DR / PROVINCIA + SELECCIÓN (UI)
# ==============================================

# -------- Detección / normalización de DR ----------
def _infer_dr_from_delegacion(name: str) -> str:
    """
    Intenta inferir la Dirección Regional (DR) desde el texto de 'Delegación'.
    Soporta patrones:
      - 'R 3', 'R3', 'r 7'
      - 'DR-1', 'DR-02'...
    """
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
    """
    Busca una columna que 'suene' a Dirección Regional.
    Retorna el nombre real de la columna o None.
    """
    norm_map = {_norm_str(c): c for c in df.columns}
    candidates = [
        "direccionregional", "direccion regional", "dirregional",
        "dr", "region", "regional", "direccion", "direccionreg", "regiones"
    ]
    for cand in candidates:
        if cand in norm_map:
            return norm_map[cand]
    # heurística: contiene 'direccion' y 'regional' o es exactamente 'dr'
    for k, real in norm_map.items():
        if ("direccion" in k and "regional" in k) or k == "dr":
            return real
    return None

def _dr_sort_key(s: str):
    """
    Key de orden para listas de DR: intenta ordenar por número (R1, R2, ...).
    Coloca no-identificados al final.
    """
    if not isinstance(s, str):
        return (999, "")
    m = re.search(r"r\s*([0-9]+)", s, flags=re.IGNORECASE)
    num = int(m.group(1)) if m else 999
    return (num, s)

# -------- Detección de columna 'Provincia' ----------
def _pick_prov_column(df: pd.DataFrame) -> Optional[str]:
    """
    Busca una columna que represente la provincia (nombre u otra variante).
    """
    norm_map = {_norm_str(c): c for c in df.columns}
    candidates = ["provincia", "province", "prov", "provincial", "provincia nombre", "nom provincia"]
    for cand in candidates:
        if cand in norm_map:
            return norm_map[cand]
    for k, real in norm_map.items():
        if "provinc" in k:
            return real
    return None

# -------- Toggle: usar total vs. selección actual ----------
def _scope_total_o_seleccion(
    df_selected: pd.DataFrame,
    df_all_options: pd.DataFrame,
    key: str,
    etiqueta_plural: str
):
    """
    Pequeño helper de UI para ofrecer 'ver todo' o 'solo la selección'.
    Devuelve (scope_df, using_total: bool).
    """
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
# =======================================================
# PARTE 7/10 — DASHBOARD: CONSOLIDADO + DR + PESTAÑAS
# =======================================================

st.divider()
st.header("📊 Dashboard de Avance (extra)")

with st.expander("ℹ️ Instrucciones", expanded=True):
    st.markdown("""
    1) **Carga** aquí el **Excel consolidado** (hoja `resumen`) generado arriba.  
    2) Navega por las pestañas: **Por Delegación**, **Por Dirección Regional**, **Gobierno Local (por Provincia)**.  
    3) En cada pestaña verás métricas de avance y paneles GL/FP.
    """)

dash_file = st.file_uploader(
    "Cargar Excel consolidado (resumen_matrices.xlsx)",
    type=["xlsx"],
    key="dash_excel"
)

if dash_file:
    # 1) Cargar hoja 'resumen' (fallback: primera hoja)
    try:
        df_dash = pd.read_excel(dash_file, sheet_name="resumen")
    except Exception:
        df_dash = pd.read_excel(dash_file)

    # 2) Asegurar numéricos en columnas clave
    df_dash = _ensure_numeric(df_dash.copy())

    # 3) DR_inferida
    dr_col = _pick_dr_column(df_dash)
    if dr_col:
        # usar la columna detectada tal cual (normalizando blancos/vacíos)
        tmp = df_dash[dr_col].astype(str)
        tmp = tmp.replace(
            {"": "Sin DR / No identificado", "nan": "Sin DR / No identificado", "none": "Sin DR / No identificado"}
        )
        df_dash["DR_inferida"] = tmp
    else:
        # inferir desde Delegación si existe
        if "Delegación" in df_dash.columns:
            df_dash["DR_inferida"] = df_dash["Delegación"].astype(str).apply(_infer_dr_from_delegacion)
        else:
            df_dash["DR_inferida"] = "Sin DR / No identificado"

    # 4) Pestañas
    tabs = st.tabs(["🏢 Por Delegación", "🗺️ Por Dirección Regional", "🏛️ Gobierno Local (por Provincia)"])

    # ======================= TAB 1: POR DELEGACIÓN =======================
    with tabs[0]:
        st.subheader("Avance por Delegación Policial")

        if "Delegación" not in df_dash.columns:
            st.info("El Excel no contiene la columna 'Delegación'.")
        else:
            delegs = sorted(df_dash["Delegación"].dropna().astype(str).unique().tolist())
            sel = st.selectbox("Delegación Policial", delegs, index=0, key="sel_deleg")

            # Alcance: selección vs total
            dsel = df_dash[df_dash["Delegación"] == sel]
            scope_df, using_total = _scope_total_o_seleccion(
                df_selected=dsel,
                df_all_options=df_dash,
                key="toggle_total_deleg",
                etiqueta_plural="delegaciones"
            )

            # Agregados numéricos
            agg = scope_df.select_dtypes(include=[np.number]).sum(numeric_only=True)

            # Total de indicadores (si no está, derivar GL+FP)
            total_ind = agg.get("Total Indicadores", np.nan)
            if np.isnan(total_ind):
                total_ind = agg.get("Indicadores Gobierno Local", 0) + agg.get("Indicadores Fuerza Pública", 0)

            # Avance global (n y %)
            sin_n  = agg.get("Sin actividades (n)", 0)
            con_n  = agg.get("Con actividades (n)", 0)
            comp_n = agg.get("Completos (n)", 0)

            def _pct(n, d): 
                return (n / d * 100.0) if d > 0 else 0.0

            sin_p, con_p, comp_p = _pct(sin_n, total_ind), _pct(con_n, total_ind), _pct(comp_n, total_ind)

            titulo_h3 = "Total (todas las delegaciones)" if using_total else sel
            st.markdown(f"<h3 style='text-align:center;margin-top:0;color:#111;'>{titulo_h3}</h3>", unsafe_allow_html=True)

            # Número grande: Líneas de Acción (si existe columna)
            # (Si tu consolidado no trae desglose, puede ser 0; no es error)
            lineas_total = int(agg.get("Líneas de Acción", 0))
            _big_number(lineas_total, "Líneas de Acción", big_px=72)

            # Gráfico de avance total (GL+FP)
            _bar_avance((sin_p, con_p, comp_p), title="Total de indicadores (Gobierno Local + Fuerza Pública)")

            # Paneles GL / FP
            top_gl, top_fp = st.columns(2)

            gl_tot = agg.get("Indicadores Gobierno Local", 0)
            gl_sin_n  = agg.get("GL Sin actividades (n)", 0)
            gl_con_n  = agg.get("GL Con actividades (n)", 0)
            gl_comp_n = agg.get("GL Completos (n)", 0)
            gl_sin_p  = _pct(gl_sin_n, gl_tot)
            gl_con_p  = _pct(gl_con_n, gl_tot)
            gl_comp_p = _pct(gl_comp_n, gl_tot)
            _panel_tres(top_gl, "Gobierno Local", gl_sin_n, gl_sin_p, gl_con_n, gl_con_p, gl_comp_n, gl_comp_p, gl_tot)

            fp_tot = agg.get("Indicadores Fuerza Pública", 0)
            fp_sin_n  = agg.get("FP Sin actividades (n)", 0)
            fp_con_n  = agg.get("FP Con actividades (n)", 0)
            fp_comp_n = agg.get("FP Completos (n)", 0)
            fp_sin_p  = _pct(fp_sin_n, fp_tot)
            fp_con_p  = _pct(fp_con_n, fp_tot)
            fp_comp_p = _pct(fp_comp_n, fp_tot)
            _panel_tres(top_fp, "Fuerza Pública", fp_sin_n, fp_sin_p, fp_con_n, fp_con_p, fp_comp_n, fp_comp_p, fp_tot)

            # Resumen global debajo
            bottom = st.container()
            _resumen_avance(bottom, sin_n, sin_p, con_n, con_p, comp_n, comp_p, total_ind)

    # =================== TAB 2: POR DIRECCIÓN REGIONAL ===================
    with tabs[1]:
        st.subheader("Avance por Dirección Regional (DR)")

        drs = sorted(df_dash["DR_inferida"].astype(str).unique().tolist(), key=_dr_sort_key)
        # elegir primer DR identificado por defecto
        idx_default = next((i for i, v in enumerate(drs) if v and "sin dr" not in v.lower()), 0)
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

            lineas_total = int(agg.get("Líneas de Acción", 0))
            _big_number(lineas_total, "Líneas de Acción", big_px=72)
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

                def _pct(n, d):
                    return (n / d * 100.0) if d > 0 else 0.0

                # Solo GL en este tab
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

                lineas_total = int(agg.get("Líneas de Acción", 0))
                _big_number(lineas_total, "Líneas de Acción", big_px=72)
                _bar_avance((gl_sin_p, gl_con_p, gl_comp_p), title="Total de indicadores (Gobierno Local)")

                _panel_tres(st.container(), "Gobierno Local",
                            gl_sin_n, gl_sin_p, gl_con_n, gl_con_p, gl_comp_n, gl_comp_p, gl_tot)
else:
    st.info("Carga el Excel consolidado para habilitar los dashboards.")
# ===================================================
# PARTE 8/10 — HELPERS: LÍNEAS Y BANDERAS POR FILA
# ===================================================
# Estas utilidades se usan en el resumen por categorías (Partes 9–10).

def _sum_lineas_df(df_any: pd.DataFrame) -> Tuple[int, int, int, int]:
    """
    Devuelve (total, gl, fp, mx) para un subconjunto del consolidado.
    Prioriza columnas de desglose de Líneas de Acción si existen.
    Si no existen, intenta 'Líneas de Acción' total.
    Si tampoco, hace fallback a contar filas (último recurso).
    """
    if df_any is None or df_any.empty:
        return 0, 0, 0, 0

    # Solo columnas numéricas
    num = df_any.select_dtypes(include=[np.number])

    cols_break = {
        "gl": "Líneas de Acción Gobierno Local",
        "fp": "Líneas de Acción Fuerza Pública",
        "mx": "Líneas de Acción Mixtas",
    }

    if set(cols_break.values()).issubset(set(num.columns)):
        gl = int(num[cols_break["gl"]].sum())
        fp = int(num[cols_break["fp"]].sum())
        mx = int(num[cols_break["mx"]].sum())
        return gl + fp + mx, gl, fp, mx

    if "Líneas de Acción" in num.columns:
        tot = int(num["Líneas de Acción"].sum())
        return tot, 0, 0, 0

    # Último recurso: contar filas (proxy de cantidad de líneas registradas)
    return len(df_any), 0, 0, 0


def _role_flags_from_row(row: pd.Series) -> Tuple[bool, bool, bool]:
    """
    Determina si una fila del consolidado está atendida por:
      - Gobierno Local, Fuerza Pública, Mixta.
    Regla:
      - Si hay columnas de 'Líneas de Acción {GL/FP/Mixtas}', usamos >0.
      - Si no hay, usamos 'Indicadores Gobierno Local' / 'Indicadores Fuerza Pública' >0.
      - 'Mixta' es True si (GL y FP) cuando no exista columna explícita de Mixtas.
    """
    la_gl = row.get("Líneas de Acción Gobierno Local", np.nan)
    la_fp = row.get("Líneas de Acción Fuerza Pública", np.nan)
    la_mx = row.get("Líneas de Acción Mixtas", np.nan)

    ind_gl = row.get("Indicadores Gobierno Local", np.nan)
    ind_fp = row.get("Indicadores Fuerza Pública", np.nan)

    has_la_cols = not (pd.isna(la_gl) and pd.isna(la_fp) and pd.isna(la_mx))

    if has_la_cols:
        gl = _to_num_safe(la_gl) > 0 if not pd.isna(la_gl) else False
        fp = _to_num_safe(la_fp) > 0 if not pd.isna(la_fp) else False
        mx = _to_num_safe(la_mx) > 0 if not pd.isna(la_mx) else False
        # Si no hay registro explícito de mixtas y ambos lados están presentes, consideramos Mixta
        if not mx and (gl and fp):
            mx = True
    else:
        gl = _to_num_safe(ind_gl) > 0 if not pd.isna(ind_gl) else False
        fp = _to_num_safe(ind_fp) > 0 if not pd.isna(ind_fp) else False
        mx = gl and fp

    return gl, fp, mx


def _lineas_tot_y_desglose_from_agg(agg: pd.Series) -> Tuple[int, Optional[int], Optional[int], Optional[int], bool]:
    """
    A partir de una serie agregada (sumas de numeric_only=True),
    retorna: (total, gl, fp, mx, has_breakdown)
    - Si existen columnas de desglose de Líneas de Acción, devuelve cada una.
    - Si no, devuelve solo el total y 'has_breakdown=False'.
    """
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
# ==========================================================
# PARTE 9/10 — RESUMEN LIMPIO (solo atendidos, en viñetas)
# ==========================================================
# - Auto-detecta columnas de Categorías y Problemática
# - Muestra SOLO problemáticas con atención (GL/FP/Mixta)
# - Sin tablas HTML y SIN "(GL:x, FP:y, Mixtas:z)"

def _autodetect_found_cols(df_any: pd.DataFrame) -> Dict[str, str]:
    """
    Busca automáticamente nombres de columnas compatibles con:
      - Categorías (categorias/categoría/categoria/categorías)
      - Problemática (problematica/problemática/problema/problemas)
    Devuelve un dict con los nombres reales encontrados.
    """
    return {
        k: v for k, v in {
            "Categorias": _pick_named_column(df_any, ["categorias", "categoria", "categoría", "categorías"]),
            "Problematica": _pick_named_column(df_any, ["problematica", "problemática", "problema", "problemas"]),
        }.items() if v is not None
    }


def _render_segundo_resumen_limpio(df_filt: pd.DataFrame, found_cols: Dict[str, str]):
    """
    Render del resumen en formato claro:
      **Se atienden N líneas de acción.**
      <Categoría>
      • Problema — atendido por: Gobierno Local, Fuerza Pública, Mixta

    Reglas:
    - SOLO listar problemáticas con atención (GL/FP/Mixta).
    - No se muestra desglose "(GL:x, FP:y, Mixtas:z)" en el título.
    - Orden sugerido de categorías: primero 'Riesgo(s)', luego 'Delito(s)', luego alfabético.
    """
    st.markdown("---")
    st.markdown("### 🧾 Resumen por categorías (solo atendidos)")

    if df_filt is None or df_filt.empty:
        st.info("No hay registros para los datos seleccionados.")
        return

    col_categorias = found_cols.get("Categorias")
    col_problema   = found_cols.get("Problematica")
    if not col_categorias or not col_problema:
        st.caption("No se encontraron columnas de **Categorías** y **Problemática** para resumir.")
        return

    # Total de líneas de acción (sin desglose en el texto)
    total_lineas, _, _, _ = _sum_lineas_df(df_filt)
    st.markdown(f"**Se atienden {int(total_lineas)} líneas de acción.**")

    # Construir: categoría → lista de viñetas (solo si la fila tiene algún rol)
    cat_map: Dict[str, List[str]] = {}

    for _, row in df_filt.iterrows():
        gl_f, fp_f, mx_f = _role_flags_from_row(row)
        if not (gl_f or fp_f or mx_f):
            continue  # omitimos filas sin atención

        roles = []
        if gl_f: roles.append("Gobierno Local")
        if fp_f: roles.append("Fuerza Pública")
        if mx_f: roles.append("Mixta")

        cats = _extract_tokens(row[col_categorias])
        probs = _extract_tokens(row[col_problema])

        for cat in cats:
            for prob in probs:
                cat_map.setdefault(cat, []).append(f"• {prob} — atendido por: {', '.join(roles)}")

    # Orden de categorías: Riesgo(s) -> Delito(s) -> alfabético
    def _cat_order_key(s: str):
        n = _norm_str(s)
        if "riesgo" in n:
            return (0, s.lower())
        if "delito" in n:
            return (1, s.lower())
        return (2, s.lower())

    printed_any = False
    for cat in sorted(cat_map.keys(), key=_cat_order_key):
        items = cat_map[cat]
        if not items:
            continue
        printed_any = True
        st.markdown(f"**{cat}**")
        st.markdown("\n".join(items))

    if not printed_any:
        st.caption("No hay problemáticas atendidas en las categorías disponibles.")
