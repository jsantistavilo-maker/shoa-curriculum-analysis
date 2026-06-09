"""
Módulo de análisis comparativo curricular.
Calcula deltas, porcentajes, clasificaciones y métricas de criticidad.
"""

import pandas as pd
import numpy as np
from data_loader import CURRICULA, CURRICULA_LABELS

# Currículos internacionales (excluye SHOA que es la referencia)
INTL = ["padilla", "sweden", "uss", "ucl"]

# Umbral para clasificación (±15 %)
THRESHOLD_PCT = 15.0


# ---------------------------------------------------------------------------
# Función principal de análisis
# ---------------------------------------------------------------------------

def analyze(df: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe df_subtopics (o df_topics) y devuelve un DataFrame enriquecido con:

    - shoa, padilla, sweden, uss, ucl          (horas absolutas)
    - pct_shoa, pct_padilla, …                  (% del total de cada currículo)
    - intl_avg                                  (promedio de los 4 internacionales)
    - intl_median                               (mediana internacional)
    - delta_padilla, delta_sweden, delta_uss, delta_ucl  (SHOA – intl, en horas)
    - delta_avg                                 (SHOA – promedio_intl)
    - delta_avg_pct                             (delta_avg relativo al promedio intl)
    - clasificacion                             (SOBREVALORADO / SUBVALORADO / ALINEADO)
    - urgencia                                  (ALTA / MEDIA / — )
    - criticidad                                (índice de criticidad)
    - n_curricula_mayor                         (# curriculos donde SHOA > ellos)
    """
    df = df.copy()

    # Totales de cada currículo (para calcular % del total)
    totals = {c: df[c].sum() for c in CURRICULA}
    totals = {c: max(t, 1) for c, t in totals.items()}   # evitar div/0

    # Porcentajes dentro de cada currículo
    for c in CURRICULA:
        df[f"pct_{c}"] = (df[c] / totals[c] * 100).round(2)

    # Promedio y mediana internacionales
    df["intl_avg"]    = df[INTL].mean(axis=1).round(2)
    df["intl_median"] = df[INTL].median(axis=1).round(2)

    # Deltas individuales (SHOA – intl, positivo = SHOA tiene MÁS)
    for c in INTL:
        df[f"delta_{c}"] = (df["shoa"] - df[c]).round(2)

    # Delta vs promedio internacional
    df["delta_avg"] = (df["shoa"] - df["intl_avg"]).round(2)

    # Delta como % del promedio internacional
    # Caso especial: si intl_avg == 0 y SHOA > 0 → sobrevalorado extremo (100%)
    def _delta_pct(r):
        if r["intl_avg"] > 0:
            return r["delta_avg"] / r["intl_avg"] * 100
        if r["shoa"] > 0:
            return 100.0    # SHOA tiene horas, ningún otro currículo las tiene
        return 0.0          # Todos en cero

    df["delta_avg_pct"] = df.apply(_delta_pct, axis=1).round(2)

    # Número de curriculos internacionales donde SHOA > ellos
    df["n_curricula_mayor"] = df.apply(
        lambda r: sum(r["shoa"] > r[c] for c in INTL), axis=1
    )

    # Índice de criticidad: exceso_horas × n_curriculos_con_menor_carga
    df["criticidad"] = df.apply(
        lambda r: max(0.0, r["delta_avg"]) * r["n_curricula_mayor"], axis=1
    ).round(2)

    # Clasificación
    def _clasify(row):
        p = row["delta_avg_pct"]
        if p > THRESHOLD_PCT:
            return "SOBREVALORADO"
        if p < -THRESHOLD_PCT:
            return "SUBVALORADO"
        return "ALINEADO"

    df["clasificacion"] = df.apply(_clasify, axis=1)

    # Urgencia (solo para sobrevalorados)
    def _urgency(row):
        if row["clasificacion"] != "SOBREVALORADO":
            return "—"
        if row["delta_avg_pct"] > 30:
            return "ALTA"
        return "MEDIA"

    df["urgencia"] = df.apply(_urgency, axis=1)

    return df


# ---------------------------------------------------------------------------
# KPIs globales
# ---------------------------------------------------------------------------

def compute_kpis(df_analyzed: pd.DataFrame) -> dict:
    """Calcula los KPIs de resumen para el Tab 1 del dashboard."""
    total_hours = {c: df_analyzed[c].sum() for c in CURRICULA}

    counts = df_analyzed["clasificacion"].value_counts().to_dict()
    n_sobre  = counts.get("SOBREVALORADO", 0)
    n_sub    = counts.get("SUBVALORADO", 0)
    n_alin   = counts.get("ALINEADO", 0)
    n_total  = len(df_analyzed)
    convergence = round(n_alin / n_total * 100, 1) if n_total > 0 else 0.0

    # Top 3 módulos más críticos
    top3 = (
        df_analyzed[df_analyzed["clasificacion"] == "SOBREVALORADO"]
        .nlargest(3, "criticidad")[["nombre", "shoa", "intl_avg", "delta_avg", "criticidad"]]
    )

    return {
        "total_hours":  total_hours,
        "n_sobre":      n_sobre,
        "n_sub":        n_sub,
        "n_alin":       n_alin,
        "n_total":      n_total,
        "convergence":  convergence,
        "top3":         top3,
    }


# ---------------------------------------------------------------------------
# Horas sugeridas para módulos desalineados
# ---------------------------------------------------------------------------

def suggested_hours(row: pd.Series) -> float:
    """
    Calcula las horas sugeridas para un módulo:
    - Sobrevalorado: apuntar a la mediana internacional
    - Subvalorado:   apuntar al promedio internacional
    - Alineado:      mantener horas actuales
    """
    if row["clasificacion"] == "SOBREVALORADO":
        return round(row["intl_median"], 1)
    if row["clasificacion"] == "SUBVALORADO":
        return round(row["intl_avg"], 1)
    return round(row["shoa"], 1)


def best_reference(row: pd.Series) -> str:
    """
    Para un módulo sobrevalorado/subvalorado, identifica qué currículo
    internacional sirve mejor como referencia (el más cercano a la mediana).
    """
    if row["clasificacion"] == "ALINEADO":
        return "—"
    target = row["intl_median"] if row["clasificacion"] == "SOBREVALORADO" else row["intl_avg"]
    diffs = {c: abs(row[c] - target) for c in INTL}
    best = min(diffs, key=diffs.get)
    return CURRICULA_LABELS.get(best, best)
