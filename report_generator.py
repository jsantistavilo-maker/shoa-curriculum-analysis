"""
Generador de Resumen Ejecutivo PDF — Análisis Curricular SHOA.

Uso desde consola : python report_generator.py
Uso desde Streamlit: from report_generator import generate_pdf
                      pdf_bytes = generate_pdf(data)
"""

from __future__ import annotations

import io
import json
import sys
import time
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd

# ── matplotlib (non-interactive) ─────────────────────────────────────────
import matplotlib
matplotlib.use("Agg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg

# ── ReportLab ────────────────────────────────────────────────────────────
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, NextPageTemplate,
    Paragraph, Spacer, Table, TableStyle, Image,
    PageBreak, HRFlowable, KeepTogether,
)
from reportlab.pdfgen import canvas as rl_canvas

# ── Módulos locales ───────────────────────────────────────────────────────
from analysis import analyze, compute_kpis, INTL
from data_loader import CURRICULA, CURRICULA_LABELS
from intervention_analysis import (
    build_topic_assignment_map, compute_intervention_analysis,
    compute_intervention_kpis,
)

# ── Colores institucionales ───────────────────────────────────────────────
C_NAVY   = colors.HexColor("#003366")
C_BLUE   = colors.HexColor("#0077C8")
C_GOLD   = colors.HexColor("#C8A84B")
C_GREEN  = colors.HexColor("#1A7A4A")
C_PURPLE = colors.HexColor("#6C3483")
C_TEXT   = colors.HexColor("#2C3E50")
C_GRAY   = colors.HexColor("#F2F4F7")
C_CRITICA= colors.HexColor("#FADBD8")
C_ALTA   = colors.HexColor("#FAE5D3")
C_ALIN   = colors.HexColor("#D5F5E3")
C_SUB    = colors.HexColor("#D6EAF8")

MPL_COLORS = ["#003366","#0077C8","#C8A84B","#1A7A4A","#6C3483"]
CLASIF_FILL= {"CRÍTICA": C_CRITICA,"ALTA": C_ALTA,"ALINEADA": C_ALIN,"SUBESTIMADA": C_SUB}

PAGE_W, PAGE_H = A4
MARGIN = 2 * cm


# ════════════════════════════════════════════════════════════════════════════
# Canvas con numeración "Página X de Y"
# ════════════════════════════════════════════════════════════════════════════

class _NumberedCanvas(rl_canvas.Canvas):
    def __init__(self, *args, logo_path=None, **kwargs):
        super().__init__(*args, **kwargs)
        self._pages: list[dict] = []
        self._logo_path = logo_path

    def showPage(self):
        self._pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        n = len(self._pages)
        for i, state in enumerate(self._pages, start=1):
            self.__dict__.update(state)
            if i > 1:                       # portada sin header/footer
                self._draw_header_footer(i, n)
            super().showPage()
        super().save()

    def _draw_header_footer(self, page_num: int, total: int):
        self.saveState()
        # Línea dorada superior
        self.setStrokeColor(C_GOLD)
        self.setLineWidth(1.5)
        self.line(MARGIN, PAGE_H - 1.5*cm, PAGE_W - MARGIN, PAGE_H - 1.5*cm)
        # Logo pequeño en header (esquina superior izquierda)
        if self._logo_path and Path(self._logo_path).exists():
            try:
                self.drawImage(
                    self._logo_path,
                    MARGIN, PAGE_H - 1.4*cm,
                    width=1.0*cm, height=1.0*cm,
                    preserveAspectRatio=True, mask="auto",
                )
            except Exception:
                pass
        # "SHOA" en header si no hay logo
        else:
            self.setFont("Helvetica-Bold", 7)
            self.setFillColor(C_NAVY)
            self.drawString(MARGIN, PAGE_H - 1.2*cm, "SHOA")
        # Texto derecho header
        self.setFont("Helvetica", 7)
        self.setFillColor(C_TEXT)
        self.drawRightString(
            PAGE_W - MARGIN, PAGE_H - 1.2*cm,
            "Análisis Curricular Comparativo"
        )
        # Línea dorada inferior
        self.setStrokeColor(C_GOLD)
        self.line(MARGIN, 1.4*cm, PAGE_W - MARGIN, 1.4*cm)
        # Número de página
        self.setFont("Helvetica", 7)
        self.setFillColor(C_TEXT)
        self.drawCentredString(
            PAGE_W / 2, 0.8*cm,
            f"Página {page_num} de {total}  |  SHOA © {date.today().year}  |  "
            "Servicio Hidrográfico y Oceanográfico de la Armada de Chile"
        )
        self.restoreState()


# ════════════════════════════════════════════════════════════════════════════
# Estilos ReportLab
# ════════════════════════════════════════════════════════════════════════════

def _styles() -> dict:
    base = {
        "title":    ParagraphStyle("title",    fontName="Helvetica-Bold",
                                   fontSize=16, textColor=C_NAVY,
                                   spaceAfter=8,  alignment=TA_LEFT),
        "subtitle": ParagraphStyle("subtitle", fontName="Helvetica-Bold",
                                   fontSize=12, textColor=C_NAVY,
                                   spaceAfter=6,  alignment=TA_LEFT),
        "body":     ParagraphStyle("body",     fontName="Helvetica",
                                   fontSize=10, textColor=C_TEXT,
                                   spaceAfter=8,  leading=14,
                                   alignment=TA_JUSTIFY),
        "small":    ParagraphStyle("small",    fontName="Helvetica",
                                   fontSize=8,  textColor=C_TEXT,
                                   spaceAfter=4),
        "center":   ParagraphStyle("center",   fontName="Helvetica",
                                   fontSize=10, textColor=C_TEXT,
                                   alignment=TA_CENTER),
        "cover_title": ParagraphStyle("cover_title", fontName="Helvetica-Bold",
                                      fontSize=28, textColor=C_NAVY,
                                      spaceAfter=10, alignment=TA_CENTER),
        "cover_sub":   ParagraphStyle("cover_sub",   fontName="Helvetica-Bold",
                                      fontSize=14, textColor=C_NAVY,
                                      spaceAfter=8,  alignment=TA_CENTER),
        "cover_body":  ParagraphStyle("cover_body",  fontName="Helvetica",
                                      fontSize=11, textColor=C_TEXT,
                                      spaceAfter=6,  alignment=TA_CENTER),
    }
    return base


def _gold_line() -> HRFlowable:
    return HRFlowable(width="100%", thickness=1.5,
                      color=C_GOLD, spaceAfter=8, spaceBefore=4)


def _tbl_style(data_rows: int, clasif_col: int | None = None,
               clf_vals: list | None = None) -> TableStyle:
    cmds = [
        ("BACKGROUND",  (0, 0), (-1, 0),  C_NAVY),
        ("TEXTCOLOR",   (0, 0), (-1, 0),  colors.white),
        ("FONTNAME",    (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, -1), 8),
        ("GRID",        (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, C_GRAY]),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",  (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",(0,0), (-1, -1), 3),
    ]
    if clasif_col is not None and clf_vals:
        for i, v in enumerate(clf_vals, start=1):
            fill = CLASIF_FILL.get(v)
            if fill:
                cmds.append(("BACKGROUND", (0, i), (-1, i), fill))
    return TableStyle(cmds)


# ════════════════════════════════════════════════════════════════════════════
# Gráficos Matplotlib → BytesIO
# ════════════════════════════════════════════════════════════════════════════

def _fig_bytes(fig: Figure) -> io.BytesIO:
    buf = io.BytesIO()
    FigureCanvasAgg(fig).print_png(buf)
    buf.seek(0)
    return buf


def _chart_hours_comparison(totals: dict) -> io.BytesIO:
    """Barras agrupadas: horas totales por currículo."""
    fig = Figure(figsize=(6.5, 3.2), dpi=150)
    ax  = fig.add_subplot(111)
    labels = list(totals.keys())
    values = list(totals.values())
    bars = ax.bar(labels, values, color=MPL_COLORS, width=0.55, edgecolor="white")
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5,
                f"{val:.0f}", ha="center", va="bottom", fontsize=8, fontweight="bold")
    ax.set_ylabel("Horas totales", fontsize=9)
    ax.set_title("Comparativa de Horas Totales por Currículo", fontsize=10,
                 fontweight="bold", color="#003366")
    ax.spines[["top","right"]].set_visible(False)
    ax.set_facecolor("#F9F9F9")
    fig.patch.set_facecolor("white")
    fig.tight_layout(pad=1.2)
    return _fig_bytes(fig)


def _chart_top10_stacked(df_intv: pd.DataFrame) -> io.BytesIO:
    """Barras horizontales apiladas T/P/SG — top 10 asignaturas por exceso."""
    top10 = df_intv[df_intv["clasificacion"].isin(["CRÍTICA","ALTA"])].head(10)
    if top10.empty:
        top10 = df_intv.head(10)
    top10 = top10.sort_values("shoa_total", ascending=True)

    fig = Figure(figsize=(6.5, max(2.5, 0.45 * len(top10) + 0.8)), dpi=150)
    ax  = fig.add_subplot(111)
    labels = [f"{r['codigo_asig']}" for _, r in top10.iterrows()]
    T_vals  = top10["shoa_T"].tolist()
    P_vals  = top10["shoa_P"].tolist()
    SG_vals = top10["shoa_SG"].tolist()
    intl    = top10["intl_avg"].tolist()

    y = range(len(labels))
    ax.barh(list(y), T_vals,              color="#0077C8", label="T — Teóricas")
    ax.barh(list(y), P_vals,  left=T_vals, color="#1A7A4A", label="P — Prácticas")
    ax.barh(list(y), SG_vals,
            left=[t+p for t,p in zip(T_vals,P_vals)],
            color="#C8A84B", label="SG — Auto estudio")
    ax.scatter(intl, list(y), color="#C0392B", zorder=5, s=40,
               marker="|", linewidths=2, label="Prom. Internacional")
    ax.set_yticks(list(y)); ax.set_yticklabels(labels, fontsize=8)
    ax.set_xlabel("Horas", fontsize=8)
    ax.set_title("Top 10 Asignaturas — Horas T/P/SG vs Promedio Internacional",
                 fontsize=9, fontweight="bold", color="#003366")
    ax.legend(loc="lower right", fontsize=7, ncol=2)
    ax.spines[["top","right"]].set_visible(False)
    ax.set_facecolor("#F9F9F9")
    fig.patch.set_facecolor("white")
    fig.tight_layout(pad=1.0)
    return _fig_bytes(fig)


def _chart_tpsg_donut(t: float, p: float, sg: float, title: str) -> io.BytesIO:
    """Gráfico de dona T/P/SG."""
    total = t + p + sg or 1
    fig = Figure(figsize=(3.2, 3.0), dpi=150)
    ax  = fig.add_subplot(111)
    wedges, texts, autotexts = ax.pie(
        [t, p, sg],
        labels=[f"T\n{t/total*100:.1f}%", f"P\n{p/total*100:.1f}%",
                f"SG\n{sg/total*100:.1f}%"],
        colors=["#0077C8","#1A7A4A","#C8A84B"],
        autopct="",
        startangle=90,
        wedgeprops=dict(width=0.5, edgecolor="white"),
        textprops=dict(fontsize=8),
    )
    ax.set_title(title, fontsize=9, fontweight="bold", color="#003366", pad=8)
    fig.patch.set_facecolor("white")
    fig.tight_layout(pad=0.5)
    return _fig_bytes(fig)


# ════════════════════════════════════════════════════════════════════════════
# Constructores de contenido por página
# ════════════════════════════════════════════════════════════════════════════

def _page_cover(s: dict, logo_path: str | None) -> list:
    story = []
    story.append(Spacer(1, 3*cm))

    # Logo
    if logo_path and Path(logo_path).exists():
        story.append(Image(logo_path, width=4*cm, height=4*cm,
                           hAlign="CENTER"))
    else:
        story.append(Paragraph("⚓", ParagraphStyle("anc", fontName="Helvetica-Bold",
                                fontSize=48, textColor=C_NAVY, alignment=TA_CENTER)))
    story.append(Spacer(1, 0.8*cm))

    story.append(HRFlowable(width="60%", thickness=2, color=C_GOLD,
                             hAlign="CENTER", spaceAfter=12, spaceBefore=4))
    story.append(Paragraph("Análisis Curricular Comparativo", s["cover_title"]))
    story.append(Paragraph(
        "Propuesta de Ajuste basada en Estándares IHO Internacionales",
        s["cover_sub"]))
    story.append(HRFlowable(width="60%", thickness=2, color=C_GOLD,
                             hAlign="CENTER", spaceBefore=8, spaceAfter=20))

    story.append(Paragraph(
        "Servicio Hidrográfico y Oceanográfico de la Armada de Chile",
        s["cover_body"]))
    story.append(Paragraph(
        "Comparativa con: Almirante Padilla  |  Sweden  |  USS  |  UCL",
        s["cover_body"]))
    story.append(Spacer(1, 1*cm))
    story.append(Paragraph(
        f"Fecha de generación: {date.today().strftime('%d de %B de %Y').capitalize()}",
        ParagraphStyle("date", fontName="Helvetica", fontSize=10,
                       textColor=C_TEXT, alignment=TA_CENTER)))
    story.append(PageBreak())
    return story


def _page_executive_summary(s: dict, kpis: dict, kpis_i: dict,
                             df_analyzed: pd.DataFrame) -> list:
    story = [Paragraph("Resumen Ejecutivo", s["title"]), _gold_line()]

    n_curricula = len(INTL)
    n_topicos   = kpis["n_total"]
    shoa_total  = round(df_analyzed["shoa"].sum(), 0)
    intl_avg_t  = round(df_analyzed["intl_avg"].sum(), 0)

    story.append(Paragraph(
        f"El presente análisis compara el currículo de hidrografía del SHOA con "
        f"<b>{n_curricula} currículos internacionales</b> de referencia, evaluando "
        f"<b>{n_topicos} sub-tópicos IHO</b>. "
        f"Se identificaron <b>{kpis_i['n_critica']} asignaturas críticas</b> que requieren "
        f"intervención prioritaria, <b>{kpis_i['n_alta']} asignaturas de intervención alta</b>, "
        f"y un exceso total estimado de <b>{kpis_i['exceso_total']:.0f} horas</b> "
        f"respecto al promedio internacional.",
        s["body"]))

    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("Indicadores principales", s["subtitle"]))

    conv = kpis["convergence"]
    kpi_rows = [
        ["Indicador", "Valor"],
        ["Total sub-tópicos IHO analizados",        str(n_topicos)],
        ["Total asignaturas SHOA",                  str(kpis_i["n_total"])],
        ["Total horas SHOA",                        f"{shoa_total:.0f} h"],
        ["Promedio horas currículos internacionales",f"{intl_avg_t:.0f} h"],
        ["Diferencia total de horas",               f"{shoa_total - intl_avg_t:+.0f} h"],
        ["Asignaturas críticas (>30% exceso)",       str(kpis_i["n_critica"])],
        ["Asignaturas intervención alta (15–30%)",   str(kpis_i["n_alta"])],
        ["Asignaturas alineadas (±15%)",             str(kpis_i["n_alin"])],
        ["Asignaturas subestimadas (>15% déficit)",  str(kpis_i["n_sub"])],
        ["Índice de convergencia global",            f"{conv:.1f}%"],
    ]
    col_w = [(PAGE_W - 2*MARGIN)*0.72, (PAGE_W - 2*MARGIN)*0.28]
    tbl = Table(kpi_rows, colWidths=col_w, repeatRows=1)
    tbl.setStyle(_tbl_style(len(kpi_rows)-1))
    story.append(tbl)
    story.append(PageBreak())
    return story


def _page_hours_comparison(s: dict, df_analyzed: pd.DataFrame) -> list:
    story = [Paragraph("Comparativa de Horas por Currículo", s["title"]), _gold_line()]

    totals = {CURRICULA_LABELS[c]: round(df_analyzed[c].sum(), 1) for c in CURRICULA}
    shoa_t = totals[CURRICULA_LABELS["shoa"]]

    # Gráfico
    chart_buf = _chart_hours_comparison(totals)
    story.append(Image(chart_buf, width=PAGE_W - 2*MARGIN, height=7*cm))
    story.append(Spacer(1, 0.4*cm))

    # Tabla comparativa
    rows = [["Currículo", "Total Horas", "vs SHOA (%)"]]
    for key, code in zip(CURRICULA_LABELS.values(), CURRICULA):
        tot = totals[key]
        diff = f"{(tot - shoa_t)/shoa_t*100:+.1f}%" if code != "shoa" else "Base"
        rows.append([key, f"{tot:.1f} h", diff])

    col_w = [(PAGE_W - 2*MARGIN)*0.55,
             (PAGE_W - 2*MARGIN)*0.25,
             (PAGE_W - 2*MARGIN)*0.20]
    tbl = Table(rows, colWidths=col_w, repeatRows=1)
    tbl.setStyle(_tbl_style(len(rows)-1))
    story.append(tbl)
    story.append(Spacer(1, 0.5*cm))

    # Párrafo interpretativo
    sorted_intl = sorted(
        [(CURRICULA_LABELS[c], round(df_analyzed[c].sum(),1)) for c in INTL],
        key=lambda x: x[1]
    )
    mayor_dif = max(INTL, key=lambda c: abs(df_analyzed[c].sum() - shoa_t))
    intl_avg  = round(df_analyzed["intl_avg"].sum(), 1)
    mayor_men = "mayor" if shoa_t > intl_avg else "menor"
    story.append(Paragraph(
        f"SHOA presenta <b>{shoa_t:.0f} horas</b> totales, siendo <b>{mayor_men}</b> "
        f"que el promedio internacional de <b>{intl_avg:.0f} horas</b>. "
        f"La mayor diferencia se observa respecto a "
        f"<b>{CURRICULA_LABELS[mayor_dif]}</b> "
        f"({df_analyzed[mayor_dif].sum():.0f} h).",
        s["body"]))
    story.append(PageBreak())
    return story


def _page_critical_assignments(s: dict, df_intv: pd.DataFrame) -> list:
    story = [Paragraph("Asignaturas que Requieren Intervención Prioritaria",
                        s["title"]), _gold_line()]

    criticas = df_intv[df_intv["clasificacion"].isin(["CRÍTICA","ALTA"])]
    top10 = criticas.head(10)

    if not df_intv.empty:
        chart_buf = _chart_top10_stacked(df_intv)
        story.append(Image(chart_buf, width=PAGE_W - 2*MARGIN,
                           height=min(10*cm, max(4*cm, 0.6*cm*len(top10) + 1.5*cm))))

    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph(f"Top asignaturas críticas ({len(top10)} de {len(criticas)})",
                            s["subtitle"]))

    hdr = ["#", "Código", "Asignatura", "Total SHOA", "Prom. Int.", "Exceso (h)", "Exceso %", "Clasif."]
    rows = [hdr]
    clf_vals = []
    for i, (_, r) in enumerate(top10.iterrows(), start=1):
        rows.append([
            str(i), r["codigo_asig"],
            r["nombre_asig"][:30],
            f"{r['shoa_total']:.1f}",
            f"{r['intl_avg']:.1f}",
            f"+{r['delta_h']:.1f}",
            f"{r['delta_pct']:+.1f}%",
            r["clasificacion"],
        ])
        clf_vals.append(r["clasificacion"])

    col_w = [0.5*cm, 1.5*cm, 5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.8*cm]
    tbl = Table(rows, colWidths=col_w, repeatRows=1)
    tbl.setStyle(_tbl_style(len(rows)-1, clasif_col=7, clf_vals=clf_vals))
    story.append(tbl)
    story.append(PageBreak())
    return story


def _page_tpsg_analysis(s: dict, df_leaves: pd.DataFrame,
                         horas_asig: dict) -> list:
    story = [Paragraph(
        "Distribución de Horas por Tipo: Teóricas, Prácticas y Auto Estudio",
        s["title"]), _gold_line()]

    # Totales T/P/SG SHOA
    if all(c in df_leaves.columns for c in ("shoa_T","shoa_P","shoa_SG")):
        t_sh  = float(df_leaves["shoa_T"].sum())
        p_sh  = float(df_leaves["shoa_P"].sum())
        sg_sh = float(df_leaves["shoa_SG"].sum())
    elif horas_asig:
        t_sh  = sum(v["T"]  for v in horas_asig.values())
        p_sh  = sum(v["P"]  for v in horas_asig.values())
        sg_sh = sum(v["SG"] for v in horas_asig.values())
    else:
        t_sh = p_sh = sg_sh = 0.0

    total_sh = t_sh + p_sh + sg_sh or 1
    pct_T  = t_sh  / total_sh * 100
    pct_P  = p_sh  / total_sh * 100
    pct_SG = sg_sh / total_sh * 100

    # Gráficos de dona lado a lado
    dona_shoa = _chart_tpsg_donut(t_sh, p_sh, sg_sh, "SHOA")
    intl_note  = "Sin desglose T/P/SG\ndisponible"

    fig_comb = Figure(figsize=(6.5, 3.2), dpi=150)
    ax1 = fig_comb.add_subplot(121)
    ax2 = fig_comb.add_subplot(122)

    # Dona SHOA
    ax1.pie([t_sh, p_sh, sg_sh],
            labels=[f"T {pct_T:.1f}%", f"P {pct_P:.1f}%", f"SG {pct_SG:.1f}%"],
            colors=["#0077C8","#1A7A4A","#C8A84B"],
            startangle=90,
            wedgeprops=dict(width=0.5, edgecolor="white"),
            textprops=dict(fontsize=7))
    ax1.set_title("SHOA", fontsize=9, fontweight="bold", color="#003366")

    # Nota internacional
    ax2.text(0.5, 0.5,
             "Desglose T/P/SG\nno disponible\npara currículos\ninternacionales",
             ha="center", va="center", fontsize=8, color="#888888",
             transform=ax2.transAxes)
    ax2.set_facecolor("#F9F9F9")
    ax2.set_title("Internacional", fontsize=9, fontweight="bold", color="#003366")
    ax2.axis("off")
    fig_comb.patch.set_facecolor("white")
    fig_comb.tight_layout(pad=1.0)
    story.append(Image(_fig_bytes(fig_comb), width=PAGE_W - 2*MARGIN, height=6.5*cm))

    story.append(Spacer(1, 0.4*cm))

    # Tabla T/P/SG SHOA
    tbl_rows = [["Tipo de Hora", "Horas", "Porcentaje"],
                ["T — Teóricas",       f"{t_sh:.1f}",  f"{pct_T:.1f}%"],
                ["P — Prácticas",      f"{p_sh:.1f}",  f"{pct_P:.1f}%"],
                ["SG — Auto estudio",  f"{sg_sh:.1f}", f"{pct_SG:.1f}%"],
                ["Total SHOA",         f"{total_sh:.1f}", "100%"]]
    col_w = [(PAGE_W-2*MARGIN)*0.50, (PAGE_W-2*MARGIN)*0.25, (PAGE_W-2*MARGIN)*0.25]
    tbl = Table(tbl_rows, colWidths=col_w, repeatRows=1)
    tbl.setStyle(_tbl_style(len(tbl_rows)-1))
    story.append(tbl)
    story.append(Spacer(1, 0.5*cm))

    story.append(Paragraph(
        f"SHOA destina el <b>{pct_T:.1f}%</b> de sus horas a actividades teóricas, "
        f"<b>{pct_P:.1f}%</b> a prácticas y <b>{pct_SG:.1f}%</b> a auto estudio. "
        f"El perfil dominante es <b>{'T — Teóricas' if pct_T >= pct_P and pct_T >= pct_SG else 'P — Prácticas' if pct_P >= pct_SG else 'SG — Auto estudio'}</b>.",
        s["body"]))
    story.append(PageBreak())
    return story


def _page_adjustment_proposal(s: dict, df_intv: pd.DataFrame) -> list:
    story = [Paragraph("Propuesta de Ajuste Curricular", s["title"]),
             _gold_line()]
    story.append(Paragraph(
        "Basada en convergencia con estándares internacionales IHO",
        s["subtitle"]))
    story.append(Paragraph(
        "A partir del análisis comparativo se proponen los siguientes ajustes para alinear "
        "el currículo SHOA con los estándares internacionales, priorizando las asignaturas "
        "con mayor desviación respecto al promedio de los currículos de referencia.",
        s["body"]))

    criticas = df_intv[df_intv["clasificacion"].isin(["CRÍTICA","ALTA"])].copy()
    if criticas.empty:
        story.append(Paragraph("No se identificaron asignaturas para ajuste.", s["body"]))
        story.append(PageBreak())
        return story

    hdr = ["Asignatura","T act.","P act.","SG act.","Tot. act.",
           "T sug.","P sug.","SG sug.","Tot. sug.",
           "Reducción","Estrategia"]
    rows = [hdr]
    clf_vals = []
    for _, r in criticas.iterrows():
        rows.append([
            f"{r['codigo_asig']}: {r['nombre_asig'][:22]}",
            f"{r['shoa_T']:.1f}",  f"{r['shoa_P']:.1f}",  f"{r['shoa_SG']:.1f}",  f"{r['shoa_total']:.1f}",
            f"{r['sug_T']:.1f}",   f"{r['sug_P']:.1f}",   f"{r['sug_SG']:.1f}",   f"{r['sug_total']:.1f}",
            f"{r['red_total']:+.1f} h",
            r["estrategia"][:28],
        ])
        clf_vals.append(r["clasificacion"])

    # Fila de totales
    rows.append([
        "TOTALES",
        f"{criticas['shoa_T'].sum():.1f}",
        f"{criticas['shoa_P'].sum():.1f}",
        f"{criticas['shoa_SG'].sum():.1f}",
        f"{criticas['shoa_total'].sum():.1f}",
        f"{criticas['sug_T'].sum():.1f}",
        f"{criticas['sug_P'].sum():.1f}",
        f"{criticas['sug_SG'].sum():.1f}",
        f"{criticas['sug_total'].sum():.1f}",
        f"{criticas['red_total'].sum():+.1f} h",
        "",
    ])

    W = PAGE_W - 2*MARGIN
    col_w = [3.2*cm, 0.9*cm,0.9*cm,0.9*cm,1.0*cm,
             0.9*cm, 0.9*cm,0.9*cm,1.0*cm, 1.2*cm, 3.5*cm]
    tbl = Table(rows, colWidths=col_w, repeatRows=1)
    sty = _tbl_style(len(rows)-1, clasif_col=None, clf_vals=None)
    sty.add("BACKGROUND", (0, len(rows)-1), (-1, len(rows)-1),
            colors.HexColor("#FFF2CC"))
    sty.add("FONTNAME",   (0, len(rows)-1), (-1, len(rows)-1), "Helvetica-Bold")
    tbl.setStyle(sty)
    story.append(tbl)
    story.append(PageBreak())
    return story


def _page_conclusions(s: dict, df_intv: pd.DataFrame, kpis: dict,
                       kpis_i: dict, df_analyzed: pd.DataFrame) -> list:
    story = [Paragraph("Conclusiones", s["title"]), _gold_line()]

    shoa_tot  = round(df_analyzed["shoa"].sum(), 0)
    intl_avg  = round(df_analyzed["intl_avg"].sum(), 0)
    diff_h    = shoa_tot - intl_avg
    diff_pct  = diff_h / intl_avg * 100 if intl_avg else 0
    top3      = (df_intv[df_intv["clasificacion"].isin(["CRÍTICA","ALTA"])]
                 .nlargest(3, "delta_pct"))
    names_top3 = ", ".join(f"<b>{r['codigo_asig']}</b>" for _,r in top3.iterrows()) or "N/A"

    # Tipo de hora con mayor exceso
    ca = df_intv[df_intv["clasificacion"].isin(["CRÍTICA","ALTA"])]
    tipo_mayor = "T — Teóricas"
    if not ca.empty:
        ex_T  = ca["red_T"].clip(lower=0).sum()
        ex_P  = ca["red_P"].clip(lower=0).sum()
        ex_SG = ca["red_SG"].clip(lower=0).sum()
        tipo_mayor = max([("T — Teóricas", ex_T),
                          ("P — Prácticas", ex_P),
                          ("SG — Auto estudio", ex_SG)], key=lambda x: x[1])[0]

    reduccion = kpis_i["exceso_total"]
    nueva_conv = min(100, kpis["convergence"] + reduccion / shoa_tot * 100) if shoa_tot else 0

    conclusiones = [
        (f"El currículo SHOA presenta una carga horaria total de <b>{shoa_tot:.0f} horas</b>, "
         f"superando en <b>{diff_h:+.0f} horas ({diff_pct:+.1f}%)</b> al promedio internacional "
         f"de <b>{intl_avg:.0f} horas</b>."),
        (f"Se identificaron <b>{kpis_i['n_critica']} asignaturas críticas</b> (exceso >30%) y "
         f"<b>{kpis_i['n_alta']} de intervención alta</b> (exceso 15–30%), que concentran el "
         f"mayor desequilibrio respecto a los estándares IHO."),
        (f"Las asignaturas de mayor prioridad de intervención son: {names_top3}."),
        (f"El tipo de hora con mayor exceso acumulado es <b>{tipo_mayor}</b>, "
         f"lo que sugiere una revisión específica de ese componente en las asignaturas críticas."),
        (f"La implementación de los ajustes propuestos reduciría la carga total en "
         f"<b>{reduccion:.0f} horas</b>, mejorando la convergencia con el promedio "
         f"internacional al <b>{nueva_conv:.1f}%</b>."),
    ]
    for i, txt in enumerate(conclusiones, start=1):
        story.append(Paragraph(f"<b>{i}.</b> {txt}", s["body"]))

    story.append(Spacer(1, 0.5*cm))
    story.append(_gold_line())
    story.append(Paragraph("Próximos Pasos Recomendados", s["subtitle"]))
    pasos = [
        "Validación de la propuesta con equipos académicos de SHOA.",
        "Revisión detallada de asignaturas críticas identificadas.",
        "Diseño de plan de implementación gradual.",
        "Seguimiento y monitoreo post-ajuste.",
    ]
    for i, p in enumerate(pasos, start=1):
        story.append(Paragraph(f"<b>{i}.</b> {p}", s["body"]))

    story.append(PageBreak())
    return story


def _page_annex(s: dict, df_analyzed: pd.DataFrame) -> list:
    story = [Paragraph("Anexo — Detalle por Sub-Tópico IHO", s["title"]), _gold_line()]
    story.append(Paragraph(
        "Tabla completa de sub-tópicos con comparativa de horas entre currículos.",
        s["small"]))
    story.append(Spacer(1, 0.3*cm))

    hdr = ["Sub-tópico", "SHOA", "Padilla", "Sweden", "USS", "UCL", "Prom. Int.", "Δ(h)", "Clasif."]
    rows = [hdr]
    clf_vals = []
    for _, r in df_analyzed.iterrows():
        nombre = str(r["nombre"])[:32]
        rows.append([
            nombre,
            f"{r['shoa']:.1f}",
            f"{r['padilla']:.1f}",
            f"{r['sweden']:.1f}",
            f"{r['uss']:.1f}",
            f"{r['ucl']:.1f}",
            f"{r['intl_avg']:.1f}",
            f"{r['delta_avg']:+.1f}",
            r["clasificacion"],
        ])
        clf_vals.append(r["clasificacion"])

    col_w = [4.8*cm, 1.2*cm,1.2*cm,1.2*cm,1.2*cm,1.2*cm,1.2*cm,1.2*cm,2.0*cm]
    tbl = Table(rows, colWidths=col_w, repeatRows=1)
    tbl.setStyle(_tbl_style(len(rows)-1, clasif_col=8, clf_vals=clf_vals))
    story.append(tbl)
    return story


# ════════════════════════════════════════════════════════════════════════════
# Función principal
# ════════════════════════════════════════════════════════════════════════════

def generate_pdf(data: dict) -> bytes:
    """
    Genera el Resumen Ejecutivo PDF y retorna bytes.

    data: dict retornado por load_data() (incluye df_leaves, df_subtopics,
          df_topics, horas_asignaturas, etc.)
    """
    df_leaves    = data["df_leaves"]
    df_subtopics = data["df_subtopics"]
    horas_asig   = data.get("horas_asignaturas", {})

    # Análisis
    df_analyzed = analyze(df_subtopics)
    kpis        = compute_kpis(df_analyzed)

    df_exp, _   = build_topic_assignment_map(df_leaves, horas_asig)
    df_intv     = compute_intervention_analysis(df_exp) if not df_exp.empty else pd.DataFrame()
    kpis_i      = compute_intervention_kpis(df_intv)

    # Logo
    logo_path = None
    for ext in ("png", "jpg"):
        p = Path(__file__).parent / "assets" / f"logo_shoa.{ext}"
        if p.exists():
            logo_path = str(p)
            break

    # Buffer PDF
    buf = io.BytesIO()
    s   = _styles()

    # Página frame (con espacio para header/footer)
    frame_cover   = Frame(MARGIN, MARGIN, PAGE_W-2*MARGIN, PAGE_H-2*MARGIN, id="cover")
    frame_content = Frame(MARGIN, 1.8*cm, PAGE_W-2*MARGIN, PAGE_H-3.6*cm, id="content")

    def make_canvas(*args, **kwargs):
        return _NumberedCanvas(*args, logo_path=logo_path, **kwargs)

    doc = BaseDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN,  bottomMargin=MARGIN,
    )
    doc.addPageTemplates([
        PageTemplate(id="cover",   frames=[frame_cover]),
        PageTemplate(id="content", frames=[frame_content]),
    ])

    story = []
    story.extend(_page_cover(s, logo_path))
    story.append(NextPageTemplate("content"))
    story.extend(_page_executive_summary(s, kpis, kpis_i, df_analyzed))
    story.extend(_page_hours_comparison(s, df_analyzed))
    story.extend(_page_critical_assignments(s, df_intv))
    story.extend(_page_tpsg_analysis(s, df_leaves, horas_asig))
    story.extend(_page_adjustment_proposal(s, df_intv))
    story.extend(_page_conclusions(s, df_intv, kpis, kpis_i, df_analyzed))
    story.extend(_page_annex(s, df_analyzed))

    doc.build(story, canvasmaker=make_canvas)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
# Ejecución desde consola
# ════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    start = time.time()

    json_path = Path(__file__).parent / "data" / "curriculum_data.json"
    if not json_path.exists():
        print("❌ No se encontró data/curriculum_data.json — ejecuta generate_data.py primero.")
        sys.exit(1)

    print("📖 Cargando datos desde JSON...")
    with open(json_path, encoding="utf-8") as f:
        raw = json.load(f)

    from data_loader import CURRICULA
    data = {
        "df_leaves":         pd.DataFrame(raw["leaves"]),
        "df_subtopics":      pd.DataFrame(raw["subtopics"]),
        "df_topics":         pd.DataFrame(raw["topics"]),
        "horas_asignaturas": raw.get("horas_asignaturas", {}),
        "subtopic_names":    raw.get("subtopic_names", {}),
        "topic_names":       raw.get("topic_names",    {}),
        "section_names":     raw.get("section_names",  {}),
    }
    for c in CURRICULA:
        for df in (data["df_leaves"], data["df_subtopics"], data["df_topics"]):
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    for col in ("shoa_T", "shoa_P", "shoa_SG"):
        if col in data["df_leaves"].columns:
            data["df_leaves"][col] = pd.to_numeric(
                data["df_leaves"][col], errors="coerce").fillna(0.0)

    print("📄 Generando PDF...")
    pdf_bytes = generate_pdf(data)

    out_dir  = Path(__file__).parent / "outputs"
    out_dir.mkdir(exist_ok=True)
    out_name = f"SHOA_Resumen_Ejecutivo_{date.today().isoformat()}.pdf"
    out_path = out_dir / out_name
    out_path.write_bytes(pdf_bytes)

    elapsed   = time.time() - start
    n_intv    = compute_intervention_kpis(
        compute_intervention_analysis(
            build_topic_assignment_map(
                data["df_leaves"], data["horas_asignaturas"]
            )[0]
        )
    )

    print(f"✅ PDF generado: {out_path}")
    print(f"📊 Asignaturas críticas incluidas: {n_intv['n_critica']}")
    print(f"⏱️  Tiempo de generación: {elapsed:.1f} segundos")
