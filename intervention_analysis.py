"""
Análisis de intervención curricular SHOA por asignatura T/P/SG.

Flujo:
  1. build_topic_assignment_map()  → tabla expandida (tópico × asignatura)
  2. compute_intervention_analysis() → análisis por asignatura
  3. compute_intervention_kpis()   → KPIs para Tab 7
"""

from __future__ import annotations
import re
import numpy as np
import pandas as pd

_CODE_RE = re.compile(r'^([A-Za-z]+\s*\d+)')

INTL            = ["padilla", "sweden", "uss", "ucl"]
THR_CRITICA     = 30.0
THR_ALTA        = 15.0
THR_SUBESTIMADA = -15.0

INTV_COLORS_HEX = {
    "CRÍTICA":     "#C0392B",
    "ALTA":        "#E67E22",
    "ALINEADA":    "#1A7A4A",
    "SUBESTIMADA": "#2471A3",
}

INTV_FILL_HEX = {
    "CRÍTICA":     "FADBD8",
    "ALTA":        "FAE5D3",
    "ALINEADA":    "D5F5E3",
    "SUBESTIMADA": "D6EAF8",
}


# ── Utilidades ──────────────────────────────────────────────────────────────

def _extract_codes(text: str) -> list[str]:
    """Extrae códigos de asignatura separando por ';'."""
    if not isinstance(text, str) or not text.strip():
        return []
    codes = []
    for part in text.split(";"):
        m = _CODE_RE.match(part.strip())
        if m:
            code = m.group(1).strip()
            if code not in codes:
                codes.append(code)
    return codes


def _delta_pct(shoa: float, intl: float) -> float:
    if intl > 0:
        return (shoa - intl) / intl * 100
    return 100.0 if shoa > 0 else 0.0


def _classify(dp: float) -> str:
    if dp > THR_CRITICA:
        return "CRÍTICA"
    if dp > THR_ALTA:
        return "ALTA"
    if dp < THR_SUBESTIMADA:
        return "SUBESTIMADA"
    return "ALINEADA"


def _tipo_reducir(clf: str, pct_T: float, pct_P: float, pct_SG: float) -> str:
    if clf not in ("CRÍTICA", "ALTA"):
        return "—"
    if pct_T > 60:
        return "Reducir horas teóricas"
    if pct_P > 50:
        return "Reducir horas prácticas"
    if pct_SG > 30:
        return "Revisar auto estudio"
    return "Reducción proporcional T/P/SG"


def _estrategia(clf: str, pct_T: float, pct_P: float, pct_SG: float) -> str:
    if clf == "SUBESTIMADA":
        return "Redistribuir sin reducir total"
    if pct_T > 60:
        return "Reducir horas teóricas, mantener prácticas"
    if pct_P > 50:
        return "Reducir horas prácticas, mantener teóricas"
    if pct_SG > 30:
        return "Reducir auto estudio"
    return "Reducción proporcional en T/P/SG"


# ── Tarea 2 ─────────────────────────────────────────────────────────────────

def build_topic_assignment_map(
    df_leaves: pd.DataFrame,
    horas_dict: dict,
) -> tuple[pd.DataFrame, dict]:
    """
    Construye tabla expandida: una fila por (tópico IHO, asignatura SHOA).
    Retorna (df_expanded, stats).
    """
    rows = []
    stats = {"divididas": 0, "no_encontrados": set(), "filas_ok": 0}

    for _, row in df_leaves.iterrows():
        asgn_text = row.get("asgn_shoa")
        if not isinstance(asgn_text, str) or not asgn_text.strip():
            continue

        codes = _extract_codes(asgn_text)
        if not codes:
            continue

        stats["filas_ok"] += 1

        shoa_T   = float(row.get("shoa_T",  0) or 0)
        shoa_P   = float(row.get("shoa_P",  0) or 0)
        shoa_SG  = float(row.get("shoa_SG", 0) or 0)
        shoa_tot = shoa_T + shoa_P + shoa_SG
        intl_avg = float(np.mean([float(row.get(c, 0) or 0) for c in INTL]))

        fue_dividida = len(codes) > 1
        if fue_dividida:
            stats["divididas"] += 1

        # Proporciones basadas en el Total de horas de cada asignatura
        if len(codes) == 1:
            props = [1.0]
            if codes[0] not in horas_dict:
                stats["no_encontrados"].add(codes[0])
        else:
            totals = [horas_dict.get(c, {}).get("Total", 0.0) for c in codes]
            for c in codes:
                if c not in horas_dict:
                    stats["no_encontrados"].add(c)
            grand = sum(totals)
            props = ([t / grand for t in totals] if grand > 0
                     else [1.0 / len(codes)] * len(codes))

        for code, prop in zip(codes, props):
            nombre = horas_dict.get(code, {}).get("nombre", code)
            rows.append({
                "codigo_topico":   row["codigo"],
                "subtopico":       row["subtopico"],
                "topico":          row["topico"],
                "seccion":         row["seccion"],
                "desc_topico":     row.get("descripcion", ""),
                "codigo_asig":     code,
                "nombre_asig":     nombre,
                "shoa_T":          round(shoa_T   * prop, 2),
                "shoa_P":          round(shoa_P   * prop, 2),
                "shoa_SG":         round(shoa_SG  * prop, 2),
                "shoa_total":      round(shoa_tot * prop, 2),
                "intl_avg_topico": round(intl_avg * prop, 2),
                "fue_dividida":    fue_dividida,
                "proporcion_pct":  round(prop * 100, 1),
            })

    return pd.DataFrame(rows) if rows else pd.DataFrame(), stats


# ── Tarea 3 ─────────────────────────────────────────────────────────────────

def compute_intervention_analysis(df_exp: pd.DataFrame) -> pd.DataFrame:
    """Agrega por asignatura y calcula clasificación, perfil T/P/SG y recomendaciones."""
    if df_exp.empty:
        return pd.DataFrame()

    rows = []
    for code, grp in df_exp.groupby("codigo_asig"):
        nombre   = grp["nombre_asig"].iloc[0]
        shoa_T   = round(grp["shoa_T"].sum(), 1)
        shoa_P   = round(grp["shoa_P"].sum(), 1)
        shoa_SG  = round(grp["shoa_SG"].sum(), 1)
        shoa_tot = round(shoa_T + shoa_P + shoa_SG, 1)
        intl_avg = round(grp["intl_avg_topico"].sum(), 1)
        n_top    = grp["codigo_topico"].nunique()

        delta_h  = round(shoa_tot - intl_avg, 1)
        delta_p  = round(_delta_pct(shoa_tot, intl_avg), 1)
        clf      = _classify(delta_p)

        pct_T  = round(shoa_T  / shoa_tot * 100, 1) if shoa_tot > 0 else 0.0
        pct_P  = round(shoa_P  / shoa_tot * 100, 1) if shoa_tot > 0 else 0.0
        pct_SG = round(shoa_SG / shoa_tot * 100, 1) if shoa_tot > 0 else 0.0
        perfil = max([("T", pct_T), ("P", pct_P), ("SG", pct_SG)], key=lambda x: x[1])[0]

        # Horas sugeridas (mantener perfil, ajustar al total internacional)
        target = intl_avg if intl_avg > 0 else shoa_tot
        sug_T  = round(target * pct_T  / 100, 1)
        sug_P  = round(target * pct_P  / 100, 1)
        sug_SG = round(target * pct_SG / 100, 1)

        rows.append({
            "codigo_asig":   code,
            "nombre_asig":   nombre,
            "shoa_T":        shoa_T,
            "shoa_P":        shoa_P,
            "shoa_SG":       shoa_SG,
            "shoa_total":    shoa_tot,
            "pct_T":         pct_T,
            "pct_P":         pct_P,
            "pct_SG":        pct_SG,
            "perfil":        perfil,
            "intl_avg":      intl_avg,
            "delta_h":       delta_h,
            "delta_pct":     delta_p,
            "clasificacion": clf,
            "n_topicos":     n_top,
            "tipo_reducir":  _tipo_reducir(clf, pct_T, pct_P, pct_SG),
            "estrategia":    _estrategia(clf, pct_T, pct_P, pct_SG),
            "sug_T":         sug_T,
            "sug_P":         sug_P,
            "sug_SG":        sug_SG,
            "sug_total":     round(sug_T + sug_P + sug_SG, 1),
            "red_T":         round(shoa_T  - sug_T,  1),
            "red_P":         round(shoa_P  - sug_P,  1),
            "red_SG":        round(shoa_SG - sug_SG, 1),
            "red_total":     round(shoa_tot - (sug_T + sug_P + sug_SG), 1),
        })

    return (pd.DataFrame(rows)
            .sort_values("delta_pct", ascending=False)
            .reset_index(drop=True))


def compute_intervention_kpis(df_i: pd.DataFrame) -> dict:
    if df_i.empty:
        return {k: 0 for k in
                ["n_total","n_critica","n_alta","n_alin","n_sub",
                 "exceso_T","exceso_P","exceso_SG","exceso_total"]}

    counts  = df_i["clasificacion"].value_counts().to_dict()
    ca      = df_i[df_i["clasificacion"].isin(["CRÍTICA","ALTA"])]
    return {
        "n_total":      len(df_i),
        "n_critica":    counts.get("CRÍTICA",     0),
        "n_alta":       counts.get("ALTA",        0),
        "n_alin":       counts.get("ALINEADA",    0),
        "n_sub":        counts.get("SUBESTIMADA", 0),
        "exceso_T":     round(ca["red_T"].clip(lower=0).sum(),   1),
        "exceso_P":     round(ca["red_P"].clip(lower=0).sum(),   1),
        "exceso_SG":    round(ca["red_SG"].clip(lower=0).sum(),  1),
        "exceso_total": round(ca["red_total"].clip(lower=0).sum(),1),
        "top3": df_i.nlargest(3, "delta_pct")[
            ["codigo_asig","nombre_asig","shoa_total","intl_avg","delta_pct"]
        ].to_dict("records"),
    }
