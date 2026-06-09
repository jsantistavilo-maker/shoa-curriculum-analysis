"""
Módulo para cargar y procesar los datos del currículo desde el archivo Excel.
Detecta automáticamente la ruta del archivo en Desktop (con o sin OneDrive).
"""

import json
import re
import shutil
import tempfile
import numpy as np
import pandas as pd
from pathlib import Path

_JSON_PATH = Path(__file__).parent / "data" / "curriculum_data.json"

# ---------------------------------------------------------------------------
# Índices de columnas (0-based) según la estructura del Excel
# ---------------------------------------------------------------------------
SHOA_COLS    = [5, 6, 7]         # T, P, SG
PADILLA_COLS = [12, 13, 14]      # T, P, SG
SWEDEN_COLS  = [20, 21, 22, 23]  # Th, Tu, Pr, SG
USS_COLS     = [31, 32, 33, 34]  # Th, Tu, Pr, SG
UCL_COLS     = [39, 40, 41, 42]  # Th, Tu, Pr, SG

# Columnas de nombres de asignaturas (0-based): E=4, L=11, T=19, AE=30, AM=38
ASGN_COLS = {
    "shoa":    4,
    "padilla": 11,
    "sweden":  19,
    "uss":     30,
    "ucl":     38,
}

_TOTAL_KW = frozenset({"total", "subtotal", "sub-total", "suma", "sum"})

CURRICULA = ["shoa", "padilla", "sweden", "uss", "ucl"]

CURRICULA_LABELS = {
    "shoa":    "SHOA (Chile)",
    "padilla": "Padilla (Colombia)",
    "sweden":  "Sweden",
    "uss":     "USS",
    "ucl":     "UCL",
}

CURRICULA_FULL = {
    "shoa":    "SHOA – Servicio Hidrográfico y Oceanográfico de la Armada (Chile)",
    "padilla": "Escuela Naval Almirante Padilla (Colombia)",
    "sweden":  "Sweden",
    "uss":     "USS",
    "ucl":     "University College London (UCL)",
}


# ---------------------------------------------------------------------------
# Búsqueda del archivo
# ---------------------------------------------------------------------------

def _find_file() -> Path | None:
    """Busca 'Tabla resumen.xlsx' en múltiples rutas posibles."""
    candidates = [
        # OneDrive Desktop (Windows con sincronización activa)
        Path.home() / "OneDrive" / "Desktop" / "shoa resumen" / "Tabla resumen.xlsx",
        # Desktop estándar
        Path.home() / "Desktop" / "shoa resumen" / "Tabla resumen.xlsx",
        # Directorio actual del proyecto
        Path.cwd() / "Tabla resumen.xlsx",
        # Mismo directorio del script
        Path(__file__).parent / "Tabla resumen.xlsx",
    ]
    for p in candidates:
        if p.exists():
            return p
    return None


# ---------------------------------------------------------------------------
# Utilidades de conversión
# ---------------------------------------------------------------------------

def _safe_float(val) -> float:
    """Convierte un valor a float seguro, manejando NaN, None y comas decimales."""
    if val is None:
        return 0.0
    if isinstance(val, float) and np.isnan(val):
        return 0.0
    try:
        if isinstance(val, str):
            val = val.replace(",", ".").strip()
        v = float(val)
        return max(0.0, v) if not np.isnan(v) else 0.0
    except (ValueError, TypeError):
        return 0.0


def _safe_asgn(row: pd.Series, col_idx: int) -> str | None:
    """Extrae el nombre de asignatura de una columna; retorna None si vacío o es subtotal."""
    if col_idx >= len(row):
        return None
    val = row.iloc[col_idx]
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.lower() in _TOTAL_KW:
        return None
    return s


def _sum_cols(row: pd.Series, col_indices: list[int]) -> float:
    """Suma las columnas indicadas de una fila, con manejo seguro de NaN."""
    return sum(_safe_float(row.iloc[i]) for i in col_indices if i < len(row))


# ---------------------------------------------------------------------------
# Clasificadores de tipo de fila
# ---------------------------------------------------------------------------

def _is_leaf(code: str) -> bool:
    """Fila de datos: código como F1.1a, H3.2b, etc."""
    return bool(re.match(r"^[A-Z]\d+\.\d+[a-z]+$", code.strip()))


def _is_subtopic_header(code: str) -> bool:
    """Encabezado de sub-tópico: 'F1.1 Physical Geodesy'."""
    return bool(re.match(r"^[A-Z]\d+\.\d+\s+.+", code.strip()))


def _is_topic_header(code: str) -> bool:
    """Encabezado de tópico: 'F1: EARTH MODELS'."""
    return bool(re.match(r"^[A-Z]\d+:\s*.+", code.strip()))


def _is_section_header(code: str) -> bool:
    """Encabezado de sección: 'SECTION I...'."""
    return code.strip().upper().startswith("SECTION")


# ---------------------------------------------------------------------------
# Extracción de jerarquía desde el código
# ---------------------------------------------------------------------------

def _subtopic_code(leaf: str) -> str:
    """F1.1a → F1.1"""
    m = re.match(r"^([A-Z]\d+\.\d+)", leaf.strip())
    return m.group(1) if m else leaf


def _topic_code(leaf: str) -> str:
    """F1.1a → F1"""
    m = re.match(r"^([A-Z]\d+)", leaf.strip())
    return m.group(1) if m else leaf


def _section_letter(leaf: str) -> str:
    """F1.1a → F (Ciencias Fundacionales), H1.2a → H (Hidrografía)."""
    m = re.match(r"^([A-Z])", leaf.strip())
    return m.group(1) if m else "?"


# ---------------------------------------------------------------------------
# Función principal de carga
# ---------------------------------------------------------------------------

_HORAS_SHEET = "Horas_Asignaturas_SHOA"
_CODE_RE      = re.compile(r'^([A-Za-z]+\s*\d+)')


def _read_horas_asignaturas(excel_path: str) -> dict:
    """
    Lee la hoja 'Horas_Asignaturas_SHOA', consolida por código de asignatura
    y retorna {código: {nombre, T, P, SG, Total}}.
    Imprime advertencias si T+P+SG ≠ Total.
    """
    try:
        df = pd.read_excel(excel_path, sheet_name=_HORAS_SHEET,
                           header=None, engine="openpyxl")
    except Exception as exc:
        print(f"⚠️  Hoja '{_HORAS_SHEET}' no encontrada: {exc}")
        return {}

    df.columns = range(len(df.columns))
    # Saltar encabezado y filas vacías
    df = df[df[0].notna() & (df[0].astype(str).str.strip() != "Module & Content")].copy()
    df["codigo"] = df[0].astype(str).str.extract(_CODE_RE, expand=False)
    df = df[df["codigo"].notna()].copy()

    for col in [2, 3, 4]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    result: dict = {}
    for code, grp in df.groupby("codigo"):
        nombre = (grp[1].mode().iloc[0]
                  if not grp[1].mode().empty else str(grp[1].iloc[0]))
        T  = float(grp[2].sum())
        P  = float(grp[3].sum())
        SG = float(grp[4].sum())
        total = round(T + P + SG, 2)
        result[code] = {"nombre": str(nombre), "T": T, "P": P, "SG": SG, "Total": total}

    return result


def _load_from_json(path: Path) -> tuple[dict | None, str | None]:
    """Carga los datos curriculares desde el JSON pre-procesado (modo online)."""
    try:
        with open(path, encoding="utf-8") as f:
            raw = json.load(f)

        df_leaves    = pd.DataFrame(raw["leaves"])
        df_subtopics = pd.DataFrame(raw["subtopics"])
        df_topics    = pd.DataFrame(raw["topics"])

        for c in CURRICULA:
            for df in (df_leaves, df_subtopics, df_topics):
                if c in df.columns:
                    df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

        meta = raw.get("metadata", {})
        return {
            "df_leaves":          df_leaves,
            "df_subtopics":       df_subtopics,
            "df_topics":          df_topics,
            "subtopic_names":     raw.get("subtopic_names", {}),
            "topic_names":        raw.get("topic_names",    {}),
            "section_names":      raw.get("section_names",  {}),
            "horas_asignaturas":  raw.get("horas_asignaturas", {}),
            "file_path":          f"JSON v{meta.get('fecha_generacion', '?')}",
            "modo":               "json",
            "metadata":           meta,
        }, None
    except Exception as exc:
        return None, f"Error al cargar datos JSON: {exc}"


def load_data(*, force_excel: bool = False) -> tuple[dict | None, str | None]:
    """
    Carga datos curriculares desde JSON (si existe) o desde el Excel original.
    force_excel=True omite el JSON y lee el Excel directamente (para generate_data.py).
    """
    if not force_excel and _JSON_PATH.exists():
        return _load_from_json(_JSON_PATH)

    file_path = _find_file()
    if file_path is None:
        msg = (
            "**Archivo no encontrado:** `Tabla resumen.xlsx`\n\n"
            "Rutas buscadas:\n"
            "- `~/OneDrive/Desktop/shoa resumen/Tabla resumen.xlsx`\n"
            "- `~/Desktop/shoa resumen/Tabla resumen.xlsx`\n"
            "- Directorio actual del proyecto\n\n"
            "Copia el archivo a una de esas rutas y recarga la página."
        )
        return None, msg

    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            shutil.copy2(file_path, tmp.name)
            tmp_path = tmp.name
        df_raw = pd.read_excel(
            tmp_path,
            header=None,
            sheet_name=0,
            engine="openpyxl",
        )
    except Exception as exc:
        return None, f"Error al leer el archivo Excel: {exc}"

    # ------------------------------------------------------------------
    # Pasada 1: construir mapas de nombres jerárquicos
    # ------------------------------------------------------------------
    subtopic_names: dict[str, str] = {}
    topic_names:    dict[str, str] = {}
    section_names:  dict[str, str] = {
        "F": "Ciencias Fundacionales",
        "H": "Ciencias Hidrográficas",
    }

    for _, row in df_raw.iterrows():
        raw = row.iloc[0]
        if pd.isna(raw):
            continue
        code = str(raw).strip()

        if _is_subtopic_header(code):
            m = re.match(r"^([A-Z]\d+\.\d+)\s+(.+)$", code)
            if m:
                subtopic_names[m.group(1)] = code  # nombre completo incl. código

        elif _is_topic_header(code):
            m = re.match(r"^([A-Z]\d+):\s*(.+)$", code)
            if m:
                topic_names[m.group(1)] = code  # nombre completo incl. código

    # ------------------------------------------------------------------
    # Pasada 2: extraer filas hoja con horas
    # ------------------------------------------------------------------
    rows = []
    for _, row in df_raw.iterrows():
        raw = row.iloc[0]
        if pd.isna(raw):
            continue
        code = str(raw).strip()
        if not _is_leaf(code):
            continue

        subtopic = _subtopic_code(code)
        topic    = _topic_code(code)
        section  = _section_letter(code)
        desc     = str(row.iloc[1]).strip() if not pd.isna(row.iloc[1]) else code

        rows.append({
            "codigo":    code,
            "subtopico": subtopic,
            "topico":    topic,
            "seccion":   section,
            "descripcion": desc,
            "shoa":    _sum_cols(row, SHOA_COLS),
            "shoa_T":  _safe_float(row.iloc[5]) if len(row) > 5 else 0.0,
            "shoa_P":  _safe_float(row.iloc[6]) if len(row) > 6 else 0.0,
            "shoa_SG": _safe_float(row.iloc[7]) if len(row) > 7 else 0.0,
            "padilla": _sum_cols(row, PADILLA_COLS),
            "sweden":  _sum_cols(row, SWEDEN_COLS),
            "uss":     _sum_cols(row, USS_COLS),
            "ucl":     _sum_cols(row, UCL_COLS),
            "asgn_shoa":    _safe_asgn(row, ASGN_COLS["shoa"]),
            "asgn_padilla": _safe_asgn(row, ASGN_COLS["padilla"]),
            "asgn_sweden":  _safe_asgn(row, ASGN_COLS["sweden"]),
            "asgn_uss":     _safe_asgn(row, ASGN_COLS["uss"]),
            "asgn_ucl":     _safe_asgn(row, ASGN_COLS["ucl"]),
        })

    if not rows:
        return None, "No se encontraron filas de datos con el patrón esperado (ej: F1.1a)."

    df_leaves = pd.DataFrame(rows)

    # ------------------------------------------------------------------
    # Agregado por sub-tópico
    # ------------------------------------------------------------------
    df_subtopics = (
        df_leaves
        .groupby(["subtopico", "topico", "seccion"], as_index=False)[CURRICULA]
        .sum()
    )
    df_subtopics["nombre"] = df_subtopics["subtopico"].map(
        lambda c: subtopic_names.get(c, c)
    )

    # ------------------------------------------------------------------
    # Agregado por tópico principal
    # ------------------------------------------------------------------
    df_topics = (
        df_leaves
        .groupby(["topico", "seccion"], as_index=False)[CURRICULA]
        .sum()
    )
    df_topics["nombre"] = df_topics["topico"].map(
        lambda c: topic_names.get(c, c)
    )

    return {
        "df_leaves":      df_leaves,
        "df_subtopics":   df_subtopics,
        "df_topics":      df_topics,
        "subtopic_names": subtopic_names,
        "topic_names":    topic_names,
        "section_names":      section_names,
        "horas_asignaturas":  _read_horas_asignaturas(tmp_path),
        "file_path":          str(file_path),
        "modo":               "excel",
    }, None
