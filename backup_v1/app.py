"""
Dashboard de Análisis Curricular Comparativo — SHOA vs Currículos Internacionales

Ejecutar con:   streamlit run app.py
"""

import base64
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path

from datetime import date
from data_loader import CURRICULA, CURRICULA_LABELS, load_data
from analysis import analyze, compute_kpis, INTL
from recommendations import build_recommendations, export_to_excel
from priority_sheet import build_priority_excel
from validate_data import run_validation, build_validation_excel
from assignment_analysis import (
    compute_assignment_analysis, compute_assignment_kpis,
    build_assignment_excel, CLASIF_COLORS_HEX,
)

# ---------------------------------------------------------------------------
# Configuración de página
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Análisis Curricular SHOA",
    page_icon="🧭",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# CSS institucional SHOA
st.markdown("""
<style>
    .stApp { background-color: #FFFFFF; }
    .block-container { padding-top: 1rem; padding-bottom: 2rem; }
    h2, h3 { color: #003366; }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #F2F4F7;
        border-radius: 8px;
        padding: 4px;
        gap: 4px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #003366 !important;
        color: white !important;
        border-radius: 6px;
    }

    /* Métricas */
    [data-testid="metric-container"] {
        background-color: white;
        border: 1px solid #E8ECF0;
        border-top: 4px solid #003366;
        border-radius: 8px;
        padding: 16px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.06);
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #003366 !important;
    }
    [data-testid="stSidebar"] .stMarkdown,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .stRadio > label,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span {
        color: white !important;
    }
    [data-testid="stSidebar"] .stSelectbox > div > div {
        background-color: #004A8F;
        color: white;
    }

    /* Botones */
    .stButton > button {
        background-color: #003366;
        color: white;
        border: none;
        border-radius: 6px;
        font-weight: 500;
        transition: all 0.2s;
    }
    .stButton > button:hover {
        background-color: #004A8F;
        border-left: 3px solid #C8A84B;
    }

    /* Tarjetas de módulos */
    .metric-card {
        background: #FFFFFF;
        border-left: 4px solid #003366;
        border-radius: 8px;
        padding: 0.8rem 1rem;
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
        margin-bottom: 0.5rem;
    }

    /* Tags de clasificación */
    .tag-sobre { background:#FADBD8; color:#C0392B; padding:2px 8px; border-radius:4px; font-weight:bold; }
    .tag-sub   { background:#D6EAF8; color:#2471A3; padding:2px 8px; border-radius:4px; font-weight:bold; }
    .tag-alin  { background:#D5F5E3; color:#1A7A4A; padding:2px 8px; border-radius:4px; font-weight:bold; }
    .tag-alta  { background:#C0392B; color:#FFFFFF; padding:2px 8px; border-radius:4px; font-weight:bold; }
    .tag-media { background:#F39C12; color:#FFFFFF; padding:2px 8px; border-radius:4px; font-weight:bold; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Paleta de colores coherente para los gráficos
# ---------------------------------------------------------------------------
COLOR_MAP = {
    "shoa":    "#003366",  # azul marino oscuro
    "padilla": "#0077C8",  # azul medio
    "sweden":  "#C8A84B",  # dorado
    "uss":     "#1A7A4A",  # verde
    "ucl":     "#6C3483",  # morado
}
COLORS_LIST = list(COLOR_MAP.values())

# ---------------------------------------------------------------------------
# Logo institucional
# ---------------------------------------------------------------------------

def _logo_html(ancho: int = 150) -> str:
    """Retorna HTML embebible del logo (base64 PNG/JPG o SVG inline)."""
    for ext, mime in [("png", "image/png"), ("jpg", "image/jpeg")]:
        p = Path(f"assets/logo_shoa.{ext}")
        if p.exists():
            b64 = base64.b64encode(p.read_bytes()).decode()
            return (f'<img src="data:{mime};base64,{b64}" width="{ancho}" '
                    f'style="max-height:{int(ancho*0.6)}px;object-fit:contain;">')
    svg_p = Path("assets/logo_shoa.svg")
    if svg_p.exists():
        svg = svg_p.read_text(encoding="utf-8")
        h = int(ancho * 0.4)
        svg = svg.replace('width="200"', f'width="{ancho}"').replace('height="80"', f'height="{h}"')
        return svg
    h = int(ancho * 0.4)
    return (f'<svg width="{ancho}" height="{h}" xmlns="http://www.w3.org/2000/svg">'
            f'<rect width="{ancho}" height="{h}" fill="#003366" rx="6"/>'
            f'<text x="{ancho//2}" y="{int(h*0.42)}" font-family="Arial" '
            f'font-size="{int(ancho*0.14)}" font-weight="bold" fill="#C8A84B" '
            f'text-anchor="middle">&#9875; SHOA</text></svg>')


def mostrar_logo(ancho: int = 200) -> None:
    """Muestra el logo SHOA en la posición actual (imagen o SVG placeholder)."""
    p = Path("assets/logo_shoa.png")
    if p.exists():
        st.image(str(p), width=ancho)
    else:
        st.markdown(_logo_html(ancho), unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Carga de datos con caché
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Cargando y procesando datos…")
def get_data():
    return load_data()


# ---------------------------------------------------------------------------
# Utilidades de display
# ---------------------------------------------------------------------------

def _tag(clasificacion: str) -> str:
    tags = {
        "SOBREVALORADO": '<span class="tag-sobre">SOBREVALORADO</span>',
        "SUBVALORADO":   '<span class="tag-sub">SUBVALORADO</span>',
        "ALINEADO":      '<span class="tag-alin">ALINEADO</span>',
    }
    return tags.get(clasificacion, clasificacion)


def _color_clasif(val):
    colors = {
        "SOBREVALORADO": "background-color:#FADBD8;color:#C0392B",
        "SUBVALORADO":   "background-color:#D6EAF8;color:#2471A3",
        "ALINEADO":      "background-color:#D5F5E3;color:#1A7A4A",
    }
    return colors.get(val, "")


def _color_urgencia(val):
    colors = {
        "ALTA":  "background-color:#FADBD8;color:#C0392B;font-weight:bold",
        "MEDIA": "background-color:#FEF9E7;color:#F39C12",
    }
    return colors.get(val, "")


def _style_tabla(df: pd.DataFrame):
    """Aplica colores condicionales a columna Clasificación y Urgencia.
    Compatible con pandas < 2.0 (applymap) y >= 2.0 (map en Styler)."""
    styler = df.style
    for col, fn in [
        ("Clasificación", _color_clasif),
        ("clasificacion", _color_clasif),
        ("Urgencia",      _color_urgencia),
        ("urgencia",      _color_urgencia),
    ]:
        if col in df.columns:
            try:
                styler = styler.map(fn, subset=[col])
            except AttributeError:
                styler = styler.applymap(fn, subset=[col])
    return styler


# ---------------------------------------------------------------------------
# Gráficos reutilizables
# ---------------------------------------------------------------------------

def _bar_comparison(df: pd.DataFrame, title: str = "Comparativa de horas por módulo"):
    """Gráfico de barras agrupadas con los 5 currículos."""
    labels = [CURRICULA_LABELS[c] for c in CURRICULA]

    fig = go.Figure()
    for c, color in COLOR_MAP.items():
        fig.add_trace(go.Bar(
            name=CURRICULA_LABELS[c],
            x=df["nombre"],
            y=df[c],
            marker_color=color,
            hovertemplate="<b>%{x}</b><br>Horas: %{y:.1f}<extra></extra>",
        ))

    fig.update_layout(
        title=title,
        barmode="group",
        xaxis_tickangle=-40,
        xaxis_title="Módulo / Sub-tópico",
        yaxis_title="Horas totales",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=480,
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#F2F4F7",
        font=dict(size=11),
        margin=dict(b=120),
    )
    return fig


def _heatmap_deltas(df_a: pd.DataFrame):
    """Heatmap: filas = módulos, columnas = currículos internacionales, valores = delta."""
    delta_cols = [f"delta_{c}" for c in INTL]
    labels = [CURRICULA_LABELS[c] for c in INTL]

    z = df_a[delta_cols].values
    text_z = [[f"{v:+.1f}" for v in row] for row in z]

    fig = go.Figure(go.Heatmap(
        z=z,
        x=labels,
        y=df_a["nombre"],
        text=text_z,
        texttemplate="%{text}",
        colorscale=[
            [0.0, "#2471A3"],   # azul = SHOA tiene MENOS
            [0.5, "#FFFFFF"],   # blanco = iguales
            [1.0, "#C0392B"],   # rojo = SHOA tiene MÁS
        ],
        zmid=0,
        colorbar=dict(title="Δ horas<br>(SHOA – Intl)"),
        hovertemplate="Módulo: %{y}<br>%{x}: %{z:+.1f} h<extra></extra>",
    ))

    fig.update_layout(
        title="Mapa de calor: diferencia de horas (SHOA – Internacional)",
        xaxis_title="Currículo internacional",
        yaxis_title="Sub-tópico",
        height=max(400, 22 * len(df_a)),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#F2F4F7",
        margin=dict(l=260, r=20, t=60, b=20),
        font=dict(size=10),
    )
    return fig


def _radar_chart(df_topics: pd.DataFrame):
    """Radar/spider chart comparando los 5 currículos a nivel de tópico."""
    totals = df_topics[CURRICULA].sum()
    totals = totals.replace(0, 1)

    categories = df_topics["topico"].tolist()
    if len(categories) < 3:
        return None

    fig = go.Figure()
    for c, color in COLOR_MAP.items():
        # Normalizar como % del total de cada currículo
        values_pct = (df_topics[c] / totals[c] * 100).tolist()
        values_pct += [values_pct[0]]  # cerrar el polígono
        cats = categories + [categories[0]]

        fig.add_trace(go.Scatterpolar(
            r=values_pct,
            theta=cats,
            name=CURRICULA_LABELS[c],
            line=dict(color=color, width=2),
            fill="toself",
            opacity=0.25,
            hovertemplate="%{theta}: %{r:.1f}%<extra>" + CURRICULA_LABELS[c] + "</extra>",
        ))

    fig.update_layout(
        polar=dict(
            radialaxis=dict(visible=True, range=[0, None]),
        ),
        title="Perfil curricular comparado (% del total por tópico)",
        legend=dict(orientation="h", yanchor="bottom", y=-0.25),
        height=520,
        paper_bgcolor="#F2F4F7",
    )
    return fig


def _horizontal_bar_delta(df_a: pd.DataFrame):
    """Barras horizontales ordenadas por delta promedio."""
    df_plot = df_a.sort_values("delta_avg", ascending=True).copy()
    colors = df_plot["clasificacion"].map({
        "SOBREVALORADO": "#C0392B",
        "SUBVALORADO":   "#2471A3",
        "ALINEADO":      "#1A7A4A",
    })

    fig = go.Figure(go.Bar(
        x=df_plot["delta_avg"],
        y=df_plot["nombre"],
        orientation="h",
        marker_color=colors,
        hovertemplate="<b>%{y}</b><br>Δ vs promedio: %{x:+.1f} h<extra></extra>",
    ))
    fig.add_vline(x=0, line_dash="dash", line_color="#666666")
    fig.update_layout(
        title="Diferencia SHOA vs promedio internacional (horas)",
        xaxis_title="Δ horas (positivo = SHOA excede promedio)",
        yaxis_title="Sub-tópico",
        height=max(400, 20 * len(df_plot)),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#F2F4F7",
        margin=dict(l=260, r=20, t=60, b=40),
        font=dict(size=10),
    )
    return fig


def _treemap_excess(df_a: pd.DataFrame):
    """Treemap de módulos sobrevalorados según su exceso de horas."""
    df_over = df_a[df_a["clasificacion"] == "SOBREVALORADO"].copy()
    if df_over.empty:
        return None

    df_over["exceso_h"] = df_over["delta_avg"].clip(lower=0)
    df_over["urgencia_label"] = df_over["urgencia"]

    fig = px.treemap(
        df_over,
        path=["urgencia_label", "nombre"],
        values="exceso_h",
        color="delta_avg_pct",
        color_continuous_scale=["#FADBD8", "#C0392B"],
        title="Distribución del exceso de horas — Módulos Sobrevalorados",
        labels={"delta_avg_pct": "% exceso", "exceso_h": "Horas exceso"},
    )
    fig.update_layout(height=460, paper_bgcolor="#F2F4F7")
    return fig


# ===========================================================================
# APLICACIÓN PRINCIPAL
# ===========================================================================

def main():
    # ── Header institucional ──────────────────────────────────────────────
    logo_h = _logo_html(ancho=130)
    st.markdown(f"""
<div style="background:linear-gradient(135deg,#003366 0%,#004A8F 100%);
            padding:18px 24px 14px;border-radius:10px;
            border-bottom:4px solid #C8A84B;margin-bottom:18px;">
  <div style="display:flex;align-items:center;gap:20px;">
    <div style="flex-shrink:0;">{logo_h}</div>
    <div>
      <h1 style="color:#FFFFFF;margin:0;font-size:1.55rem;font-weight:700;letter-spacing:0.3px;">
        Análisis Curricular Comparativo</h1>
      <p style="color:#C8A84B;margin:5px 0 0;font-size:0.92rem;font-weight:600;">
        SHOA vs Currículos Internacionales IHO</p>
      <p style="color:rgba(255,255,255,0.65);margin:2px 0 0;font-size:0.75rem;">
        Servicio Hidrográfico y Oceanográfico de la Armada de Chile</p>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

    # -----------------------------------------------------------------------
    # Cargar datos
    # -----------------------------------------------------------------------
    data, error = get_data()

    if error:
        st.error(f"### Error al cargar los datos\n\n{error}")
        st.stop()

    # ── Sidebar institucional ──────────────────────────────────────────────
    logo_sb = _logo_html(ancho=110)
    st.sidebar.markdown(f"""
<div style="text-align:center;padding:14px 0 8px;">{logo_sb}</div>
<hr style="border:none;border-top:1px solid #C8A84B;margin:6px 0 12px;">
""", unsafe_allow_html=True)

    _modo_json = data.get("modo") == "json"
    if _modo_json:
        meta = data.get("metadata", {})
        st.sidebar.markdown(
            f"<p style='font-size:0.78rem;color:rgba(255,255,255,0.8);text-align:center;"
            f"margin:0;'>📊 Datos: Versión {meta.get('fecha_generacion','—')}</p>",
            unsafe_allow_html=True,
        )
    else:
        st.sidebar.markdown(
            f"<p style='font-size:0.78rem;color:rgba(255,255,255,0.8);text-align:center;"
            f"margin:0;'>✅ Fuente: Excel local</p>",
            unsafe_allow_html=True,
        )

    df_leaves    = data["df_leaves"]
    df_subtopics = data["df_subtopics"]
    df_topics    = data["df_topics"]

    # Analizar a nivel de sub-tópico
    df_analyzed  = analyze(df_subtopics)
    df_top_anal  = analyze(df_topics)
    kpis         = compute_kpis(df_analyzed)

    # Sidebar: nivel de análisis
    st.sidebar.markdown("---")
    nivel = st.sidebar.radio(
        "Nivel de agrupación",
        ["Sub-tópico (detallado)", "Tópico principal (resumen)"],
        index=0,
    )
    df_display = df_analyzed if nivel.startswith("Sub") else df_top_anal

    # Sidebar: filtro por sección
    secciones = {"Todos": None, "Ciencias Fundacionales (F)": "F", "Ciencias Hidrográficas (H)": "H"}
    sec_label = st.sidebar.selectbox("Filtrar por sección", list(secciones.keys()))
    sec_filter = secciones[sec_label]
    if sec_filter:
        df_display = df_display[df_display["seccion"] == sec_filter].copy()

    # -----------------------------------------------------------------------
    # TABS
    # -----------------------------------------------------------------------
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 Vista General / KPIs",
        "📈 Comparativa de Horas",
        "🔥 Análisis de Diferencias",
        "⚠️ Módulos Críticos",
        "💡 Recomendaciones",
        "📚 Análisis por Asignatura",
    ])

    # =======================================================================
    # TAB 1 — VISTA GENERAL / KPIs
    # =======================================================================
    with tab1:
        st.subheader("Resumen Ejecutivo")

        # Totales de horas
        st.markdown("#### Horas totales por currículo")
        cols = st.columns(5)
        for i, c in enumerate(CURRICULA):
            with cols[i]:
                st.metric(
                    label=CURRICULA_LABELS[c],
                    value=f"{kpis['total_hours'][c]:,.0f} h",
                )

        st.divider()

        # Clasificación global
        col_a, col_b, col_c, col_d = st.columns(4)
        col_a.metric("📦 Módulos totales",   kpis["n_total"])
        col_b.metric("🔴 Sobrevalorados",    kpis["n_sobre"],
                      help="SHOA tiene >15% más horas que el promedio internacional")
        col_c.metric("🔵 Subvalorados",      kpis["n_sub"],
                      help="SHOA tiene >15% menos horas que el promedio internacional")
        col_d.metric("🟢 Alineados",         kpis["n_alin"],
                      help="Diferencia dentro del ±15%")

        st.divider()

        # Índice de convergencia
        st.markdown(
            f"#### Índice de Convergencia Global: "
            f"**{kpis['convergence']:.1f}%**  _{kpis['n_alin']} de {kpis['n_total']} módulos dentro de ±15%_"
        )

        # Barra visual
        prog_html = f"""
        <div style="background:#DDDDDD;border-radius:8px;height:22px;width:100%">
          <div style="background:#003366;border-radius:8px;height:22px;width:{kpis['convergence']}%;
                      text-align:center;color:white;line-height:22px;font-size:13px">
            {kpis['convergence']:.1f}%
          </div>
        </div>
        """
        st.markdown(prog_html, unsafe_allow_html=True)

        st.divider()

        # Top 3 módulos más críticos
        st.markdown("#### 🏆 Top 3 módulos más críticos (mayor exceso en SHOA)")
        if kpis["top3"].empty:
            st.info("No hay módulos sobrevalorados.")
        else:
            for _, r in kpis["top3"].iterrows():
                st.markdown(f"""
                <div class="metric-card">
                    <b>{r['nombre']}</b><br>
                    SHOA: <b>{r['shoa']:.1f} h</b> |
                    Prom. Intl: <b>{r['intl_avg']:.1f} h</b> |
                    Exceso: <b>+{r['delta_avg']:.1f} h</b> |
                    Criticidad: <b>{r['criticidad']:.0f}</b>
                </div>
                """, unsafe_allow_html=True)

        st.divider()

        # Distribución visual
        st.markdown("#### Distribución de clasificaciones")
        df_pie = pd.DataFrame({
            "Clasificación": ["SOBREVALORADO", "SUBVALORADO", "ALINEADO"],
            "Cantidad": [kpis["n_sobre"], kpis["n_sub"], kpis["n_alin"]],
        })
        fig_pie = px.pie(
            df_pie,
            names="Clasificación",
            values="Cantidad",
            color="Clasificación",
            color_discrete_map={
                "SOBREVALORADO": "#C0392B",
                "SUBVALORADO":   "#2471A3",
                "ALINEADO":      "#1A7A4A",
            },
            hole=0.4,
        )
        fig_pie.update_layout(height=320, paper_bgcolor="#F2F4F7")
        st.plotly_chart(fig_pie, use_container_width=True)

        st.divider()

        # ---- Validación de datos ----
        st.markdown("#### 🔍 Validación de Datos")
        st.caption(
            "Compara los datos procesados por la app contra el Excel original "
            "para detectar errores de parseo, mapeo de columnas o discrepancias numéricas."
        )

        if _modo_json:
            st.info(
                "ℹ️ La validación contra el Excel original solo está disponible en "
                "modo desarrollo local (cuando la app carga desde el archivo Excel directamente)."
            )
        elif st.button("🔍 Ejecutar validación de datos", type="secondary"):
            with st.spinner("Validando datos contra el Excel original…"):
                st.session_state["_val_results"] = run_validation()

        if "_val_results" in st.session_state:
            res = st.session_state["_val_results"]

            if "error" in res:
                st.error(f"Error en la validación: {res['error']}")
            else:
                r = res["resumen"]
                st.caption(f"Última validación ejecutada: {r['timestamp']}")

                # Tabla semáforo de las 5 validaciones
                _ICON = {"OK": "✅", "ERROR": "❌", "REVISAR": "⚠️", "INFO": "ℹ️"}
                val_labels = [
                    ("val1", "Totales por currículo"),
                    ("val2", "Cantidad de módulos"),
                    ("val3", "Módulo a módulo (numérico)"),
                    ("val4", "Emparejamiento de códigos"),
                    ("val5", "Casos especiales"),
                ]
                sem_rows = []
                for key, label in val_labels:
                    v = r[key]
                    sem_rows.append({
                        "Estado": f"{_ICON.get(v['status'], '—')} {v['status']}",
                        "Validación": label,
                        "Detalle": v["detalle"],
                    })
                st.dataframe(pd.DataFrame(sem_rows), use_container_width=True, hide_index=True)

                # Detalle expandible por validación
                with st.expander("📋 Validación 1 — Totales por currículo"):
                    rows_v1 = [
                        {
                            "Currículo": info["label"],
                            "Excel (h)":  info["excel_total"],
                            "App (h)":    info["app_total"],
                            "Diferencia": info["diferencia"],
                            "Estado":     info["status"],
                        }
                        for info in res["v1_totales"].values()
                    ]
                    st.dataframe(
                        pd.DataFrame(rows_v1)
                        .style.format({"Excel (h)": "{:.2f}", "App (h)": "{:.2f}", "Diferencia": "{:.4f}"}),
                        use_container_width=True,
                        hide_index=True,
                    )

                with st.expander("📋 Validación 2 — Conteo de módulos"):
                    v2 = res["v2_conteos"]
                    c2a, c2b, c2c = st.columns(3)
                    c2a.metric("Filas en Excel",     v2["n_excel"])
                    c2b.metric("Filas procesadas",   v2["n_procesados"])
                    c2c.metric("Diferencia",         v2["diferencia"])
                    if v2["codigos_perdidos"]:
                        st.warning("Códigos no procesados: " + ", ".join(v2["codigos_perdidos"]))
                    else:
                        st.success("Todos los módulos del Excel fueron procesados correctamente.")

                with st.expander("📋 Validación 3 — Módulo a módulo (numérico)"):
                    v3 = res["v3_modulos"]
                    errores_v3 = v3[v3["Estado"] == "ERROR"]
                    if errores_v3.empty:
                        st.success("Sin discrepancias numéricas — todos los valores coinciden exactamente.")
                    else:
                        st.warning(f"{len(errores_v3)} discrepancias detectadas")
                        st.dataframe(errores_v3, use_container_width=True, hide_index=True)

                with st.expander("📋 Validación 4 — Emparejamiento de códigos (fuzzy)"):
                    v4 = res["v4_matching"]
                    if v4.empty:
                        st.info("Sin datos de emparejamiento disponibles.")
                    else:
                        issues = v4[v4["Estado"].isin(["REVISAR", "ERROR"])]
                        if issues.empty:
                            st.success(f"Todos los {len(v4)} códigos verificados correctamente.")
                        else:
                            st.warning(f"{len(issues)} caso(s) a revisar de {len(v4)} total")
                            st.dataframe(issues, use_container_width=True, hide_index=True)

                with st.expander("📋 Validación 5 — Casos especiales"):
                    v5 = res["v5_especiales"]
                    c5a, c5b, c5c, c5d = st.columns(4)
                    c5a.metric("Exclusivos SHOA",     len(v5["exclusivo_shoa"]))
                    c5b.metric("SHOA=0, Intl>0",      len(v5["shoa_zero_intl_positivo"]))
                    c5c.metric("Valores >200h",        len(v5["anomalos"]))
                    c5d.metric("Horas negativas",      len(v5["negativos"]),
                               delta="⚠️" if not v5["negativos"].empty else None)
                    if not v5["exclusivo_shoa"].empty:
                        st.markdown("**Módulos exclusivos de SHOA (0h en todos los internacionales):**")
                        st.dataframe(v5["exclusivo_shoa"], use_container_width=True, hide_index=True)
                    if not v5["shoa_zero_intl_positivo"].empty:
                        st.markdown("**SHOA = 0h pero algún internacional tiene horas:**")
                        st.dataframe(v5["shoa_zero_intl_positivo"], use_container_width=True, hide_index=True)
                    if not v5["anomalos"].empty:
                        st.markdown("**Valores anómalos (>200h en un elemento):**")
                        st.dataframe(v5["anomalos"], use_container_width=True, hide_index=True)
                    if not v5["negativos"].empty:
                        st.error("**Horas negativas detectadas:**")
                        st.dataframe(v5["negativos"], use_container_width=True, hide_index=True)

                # Botón de descarga del Excel de validación
                st.divider()
                xlsx_val = build_validation_excel(res)
                st.download_button(
                    label="⬇️ Descargar reporte de validación (Excel)",
                    data=xlsx_val,
                    file_name=f"validacion_datos_SHOA_{date.today().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    # =======================================================================
    # TAB 2 — COMPARATIVA DE HORAS
    # =======================================================================
    with tab2:
        st.subheader("Comparativa de Horas por Módulo")

        # Selector de módulo específico
        modulos_disp = ["Todos"] + df_display["nombre"].tolist()
        sel = st.selectbox("Filtrar módulo específico", modulos_disp)
        df_bar = df_display if sel == "Todos" else df_display[df_display["nombre"] == sel]

        st.plotly_chart(_bar_comparison(df_bar), use_container_width=True)

        st.divider()
        st.markdown("#### Tabla comparativa completa")

        # Preparar tabla con formato de colores
        cols_tabla = ["nombre", "shoa", "padilla", "sweden", "uss", "ucl",
                      "intl_avg", "delta_avg", "clasificacion"]
        df_tbl = df_display[cols_tabla].copy()
        df_tbl.columns = [
            "Módulo", "SHOA", "Padilla", "Sweden", "USS", "UCL",
            "Prom. Intl", "Δ vs Prom.", "Clasificación"
        ]

        st.dataframe(
            _style_tabla(df_tbl)
            .format({
                "SHOA": "{:.1f}", "Padilla": "{:.1f}", "Sweden": "{:.1f}",
                "USS": "{:.1f}", "UCL": "{:.1f}",
                "Prom. Intl": "{:.1f}", "Δ vs Prom.": "{:+.1f}",
            }),
            use_container_width=True,
            height=420,
        )

    # =======================================================================
    # TAB 3 — ANÁLISIS DE DIFERENCIAS
    # =======================================================================
    with tab3:
        st.subheader("Análisis de Diferencias")

        # Heatmap
        st.markdown("#### Mapa de calor de diferencias (SHOA – cada currículo)")
        if len(df_display) > 0:
            st.plotly_chart(_heatmap_deltas(df_display), use_container_width=True)

        st.divider()

        # Barras horizontales
        st.markdown("#### Módulos ordenados por desalineación")
        st.plotly_chart(_horizontal_bar_delta(df_display), use_container_width=True)

        st.divider()

        # Radar chart (solo a nivel de tópico para que sea legible)
        st.markdown("#### Perfil curricular comparado por tópico principal")
        radar_data = df_top_anal if sec_filter is None else df_top_anal[df_top_anal["seccion"] == sec_filter]
        fig_radar = _radar_chart(radar_data)
        if fig_radar:
            st.plotly_chart(fig_radar, use_container_width=True)
        else:
            st.info("Se necesitan al menos 3 tópicos para generar el gráfico radar.")

    # =======================================================================
    # TAB 4 — MÓDULOS CRÍTICOS SHOA
    # =======================================================================
    with tab4:
        st.subheader("Módulos Críticos — SHOA")

        df_sobre = df_display[df_display["clasificacion"] == "SOBREVALORADO"].sort_values(
            "criticidad", ascending=False
        )
        df_sub   = df_display[df_display["clasificacion"] == "SUBVALORADO"].sort_values(
            "delta_avg", ascending=True
        )

        # ---- Sobrevalorados ----
        st.markdown(f"### 🔴 Módulos Sobrevalorados ({len(df_sobre)})")
        if df_sobre.empty:
            st.success("No hay módulos sobrevalorados.")
        else:
            # Tarjetas detalladas para los top 10
            for _, r in df_sobre.head(10).iterrows():
                urgencia_tag = f'<span class="tag-alta">ALTA</span>' if r["urgencia"] == "ALTA" \
                               else f'<span class="tag-media">MEDIA</span>'
                curricula_menor = [c for c in INTL if r["shoa"] > r[c]]
                curricula_str = ", ".join(CURRICULA_LABELS[c] for c in curricula_menor)
                st.markdown(f"""
                <div class="metric-card">
                    <b>{r['nombre']}</b>  {urgencia_tag}<br>
                    <small>
                    SHOA: <b>{r['shoa']:.1f} h</b> |
                    Prom. Intl: <b>{r['intl_avg']:.1f} h</b> |
                    Exceso: <span style="color:#CC0000"><b>+{r['delta_avg']:.1f} h ({r['delta_avg_pct']:+.0f}%)</b></span> |
                    Criticidad: <b>{r['criticidad']:.0f}</b><br>
                    Curricula con menor carga: {curricula_str if curricula_str else "—"}
                    </small>
                </div>
                """, unsafe_allow_html=True)

            st.divider()

            # Treemap
            fig_tree = _treemap_excess(df_display)
            if fig_tree:
                st.plotly_chart(fig_tree, use_container_width=True)

            st.markdown("##### Tabla completa de módulos sobrevalorados")
            cols_s = ["nombre", "shoa", "intl_avg", "intl_median",
                      "delta_avg", "delta_avg_pct", "urgencia", "criticidad",
                      "padilla", "sweden", "uss", "ucl"]
            df_s = df_sobre[cols_s].copy()
            df_s.columns = ["Módulo", "SHOA (h)", "Prom. Intl", "Mediana Intl",
                             "Δ (h)", "Δ (%)", "Urgencia", "Criticidad",
                             "Padilla", "Sweden", "USS", "UCL"]
            st.dataframe(
                _style_tabla(df_s).format({
                    "SHOA (h)": "{:.1f}", "Prom. Intl": "{:.1f}", "Mediana Intl": "{:.1f}",
                    "Δ (h)": "{:+.1f}", "Δ (%)": "{:+.1f}%", "Criticidad": "{:.0f}",
                    "Padilla": "{:.1f}", "Sweden": "{:.1f}", "USS": "{:.1f}", "UCL": "{:.1f}",
                }),
                use_container_width=True,
            )

        st.divider()

        # ---- Subvalorados ----
        st.markdown(f"### 🔵 Módulos Subvalorados ({len(df_sub)})")
        if df_sub.empty:
            st.success("No hay módulos subvalorados.")
        else:
            for _, r in df_sub.iterrows():
                curricula_mayor = [c for c in INTL if r[c] > r["shoa"]]
                curricula_str = ", ".join(
                    f"{CURRICULA_LABELS[c]} ({r[c]:.1f}h)" for c in curricula_mayor
                )
                st.markdown(f"""
                <div class="metric-card" style="border-left-color:#2471A3">
                    <b>{r['nombre']}</b><br>
                    <small>
                    SHOA: <b>{r['shoa']:.1f} h</b> |
                    Prom. Intl: <b>{r['intl_avg']:.1f} h</b> |
                    Déficit: <span style="color:#2471A3"><b>{r['delta_avg']:.1f} h ({r['delta_avg_pct']:+.0f}%)</b></span><br>
                    Curricula que lo priorizan: {curricula_str if curricula_str else "—"}
                    </small>
                </div>
                """, unsafe_allow_html=True)

    # =======================================================================
    # TAB 5 — RECOMENDACIONES CURRICULARES
    # =======================================================================
    with tab5:
        st.subheader("Recomendaciones Curriculares")
        st.markdown(
            "Las recomendaciones se generan automáticamente para todos los módulos "
            "clasificados como **SOBREVALORADOS** o **SUBVALORADOS**. "
            "Los módulos **ALINEADOS** no requieren ajuste."
        )

        df_rec = build_recommendations(df_display)

        if df_rec.empty:
            st.success("🎉 Todos los módulos están alineados con los estándares internacionales.")
        else:
            # Filtros
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                fil_clasif = st.multiselect(
                    "Filtrar por clasificación",
                    ["SOBREVALORADO", "SUBVALORADO"],
                    default=["SOBREVALORADO", "SUBVALORADO"],
                )
            with col_f2:
                fil_urg = st.multiselect(
                    "Filtrar por urgencia",
                    ["ALTA", "MEDIA", "—"],
                    default=["ALTA", "MEDIA", "—"],
                )

            mask = df_rec["Clasificación"].isin(fil_clasif) & df_rec["Urgencia"].isin(fil_urg)
            df_filtered = df_rec[mask].copy()

            st.markdown(f"**{len(df_filtered)}** recomendaciones mostradas "
                        f"({df_rec['Clasificación'].value_counts().get('SOBREVALORADO', 0)} sobrevalorados, "
                        f"{df_rec['Clasificación'].value_counts().get('SUBVALORADO', 0)} subvalorados)")

            # Tabla principal
            cols_show = [
                "Módulo", "Clasificación", "Horas SHOA", "Promedio Internacional",
                "Diferencia (h)", "% Diferencia", "Urgencia",
                "Horas Sugeridas", "Ajuste (h)", "Referencia Internacional",
            ]
            st.dataframe(
                _style_tabla(df_filtered[cols_show])
                .format({
                    "Horas SHOA": "{:.1f}",
                    "Promedio Internacional": "{:.1f}",
                    "Diferencia (h)": "{:+.1f}",
                    "% Diferencia": "{:+.1f}%",
                    "Horas Sugeridas": "{:.1f}",
                    "Ajuste (h)": "{:+.1f}",
                }),
                use_container_width=True,
                height=450,
            )

            # Botón de exportación
            st.divider()
            st.markdown("#### Exportar")
            col_exp1, col_exp2 = st.columns(2)

            with col_exp1:
                xlsx_bytes = export_to_excel(df_filtered)
                st.download_button(
                    label="⬇️ Descargar recomendaciones en Excel",
                    data=xlsx_bytes,
                    file_name="recomendaciones_curriculo_SHOA.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )

            with col_exp2:
                xlsx_prior = build_priority_excel(df_display)
                fecha_hoy  = date.today().strftime("%Y%m%d")
                st.download_button(
                    label="📋 Exportar Asignaturas Prioritarias",
                    data=xlsx_prior,
                    file_name=f"SHOA_Asignaturas_Prioritarias_{fecha_hoy}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    help="Genera un Excel con las asignaturas ordenadas por nivel de prioridad de revisión",
                )

            # Resumen narrativo
            st.divider()
            st.markdown("#### Resumen narrativo")
            n_alta = (df_rec["Urgencia"] == "ALTA").sum()
            n_media = (df_rec["Urgencia"] == "MEDIA").sum()
            n_sub = (df_rec["Clasificación"] == "SUBVALORADO").sum()

            total_exceso = df_display[df_display["clasificacion"] == "SOBREVALORADO"]["delta_avg"].sum()
            total_deficit = abs(df_display[df_display["clasificacion"] == "SUBVALORADO"]["delta_avg"].sum())

            st.info(f"""
**Resumen de ajustes recomendados:**

- **{n_alta}** módulos con urgencia **ALTA** (>30% exceso) → reducción prioritaria
- **{n_media}** módulos con urgencia **MEDIA** (15–30% exceso) → reducción recomendada
- **{n_sub}** módulos **subvalorados** → incremento recomendado

Si se implementan todas las recomendaciones:
- Reducción total estimada en módulos sobrevalorados: **{total_exceso:.0f} horas**
- Incremento total estimado en módulos subvalorados: **+{total_deficit:.0f} horas**
- Balance neto estimado: **{total_deficit - total_exceso:+.0f} horas**
            """)

    # =======================================================================
    # TAB 6 — ANÁLISIS POR ASIGNATURA
    # =======================================================================
    with tab6:
        st.subheader("Análisis por Asignatura")
        st.markdown(
            "Agrupa los tópicos por **nombre de asignatura** (columna E del Excel) "
            "y compara la carga horaria de SHOA con los equivalentes internacionales "
            "en el mismo contexto temático. Umbral de clasificación: **±20%**."
        )

        df_asgn = compute_assignment_analysis(df_leaves)

        if df_asgn.empty:
            st.warning(
                "No se encontraron datos de asignaturas. "
                "Verifica que el Excel tenga nombres en las columnas E, L, T, AE y AM."
            )
        else:
            kpis_a = compute_assignment_kpis(df_asgn)

            # ── KPIs ─────────────────────────────────────────────────────────
            st.markdown("#### Resumen por Asignatura")
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("📚 Total asignaturas",   kpis_a["n_total"])
            c2.metric("🔴 Sobreestimadas",       kpis_a["n_sobre"])
            c3.metric("🔵 Subestimadas",         kpis_a["n_sub"])
            c4.metric("🟢 Alineadas",            kpis_a["n_alin"])
            c5.metric("🟣 Exclusivas SHOA",      kpis_a["n_excl"])

            if kpis_a["top_sobre"]:
                st.info(
                    f"**Mayor sobreestimación:** {kpis_a['top_sobre']['nombre']} "
                    f"(+{kpis_a['top_sobre']['delta']:.1f} h sobre el promedio internacional)"
                )
            if kpis_a["top_sub"]:
                st.info(
                    f"**Mayor subestimación:** {kpis_a['top_sub']['nombre']} "
                    f"({kpis_a['top_sub']['delta']:.1f} h por debajo del promedio internacional)"
                )

            st.divider()

            # ── Gráfico 1: barras horizontales por clasificación ──────────────
            st.markdown("#### Carga Horaria por Asignatura SHOA")

            df_sorted = df_asgn.sort_values("horas_shoa", ascending=True)

            fig_a1 = go.Figure()

            # Leyenda: una traza dummy por clasificación
            for clf, color in CLASIF_COLORS_HEX.items():
                if (df_asgn["clasificacion"] == clf).any():
                    fig_a1.add_trace(go.Bar(
                        x=[None], y=[None], orientation="h",
                        name=clf, marker_color=color, showlegend=True,
                    ))

            # Barras reales con color individual
            fig_a1.add_trace(go.Bar(
                x=df_sorted["horas_shoa"],
                y=df_sorted["asignatura_shoa"],
                orientation="h",
                marker_color=df_sorted["clasificacion"].map(CLASIF_COLORS_HEX).tolist(),
                showlegend=False,
                hovertemplate=(
                    "<b>%{y}</b><br>"
                    "Horas SHOA: %{x:.1f}<br>"
                    "Prom. Intl: %{customdata:.1f}<extra></extra>"
                ),
                customdata=df_sorted["intl_avg"],
            ))

            # Marcadores del promedio internacional por asignatura
            fig_a1.add_trace(go.Scatter(
                x=df_sorted["intl_avg"],
                y=df_sorted["asignatura_shoa"],
                mode="markers",
                name="Prom. Internacional",
                marker=dict(
                    symbol="line-ew",
                    color="#888888",
                    size=14,
                    line=dict(width=2.5, color="#888888"),
                ),
                hovertemplate="<b>%{y}</b><br>Prom. Intl: %{x:.1f} h<extra></extra>",
            ))

            fig_a1.update_layout(
                barmode="overlay",
                xaxis_title="Horas totales",
                yaxis_title="Asignatura SHOA",
                height=max(420, 30 * len(df_asgn)),
                plot_bgcolor="#FFFFFF",
                paper_bgcolor="#F2F4F7",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                margin=dict(l=260, r=20, t=50, b=40),
                font=dict(size=10),
            )
            st.plotly_chart(fig_a1, use_container_width=True)

            st.divider()

            # ── Gráfico 2: barras agrupadas por asignatura ────────────────────
            st.markdown("#### Comparativa SHOA vs Currículos Internacionales")

            opciones_a = ["Todas"] + df_asgn["asignatura_shoa"].tolist()
            sel_a = st.selectbox("Seleccionar asignatura", opciones_a, key="sel_asgn_g2")
            df_g2 = df_asgn if sel_a == "Todas" else df_asgn[df_asgn["asignatura_shoa"] == sel_a]

            fig_a2 = go.Figure()
            curricula_g2 = [
                ("SHOA (Chile)",       "horas_shoa",    "#003366"),
                ("Padilla (Colombia)", "horas_padilla", "#0077C8"),
                ("Sweden",             "horas_sweden",  "#C8A84B"),
                ("USS",                "horas_uss",     "#1A7A4A"),
                ("UCL",                "horas_ucl",     "#6C3483"),
            ]
            for label, col, color in curricula_g2:
                fig_a2.add_trace(go.Bar(
                    name=label,
                    x=df_g2["asignatura_shoa"],
                    y=df_g2[col],
                    marker_color=color,
                    hovertemplate=f"<b>%{{x}}</b><br>{label}: %{{y:.1f}} h<extra></extra>",
                ))

            fig_a2.update_layout(
                barmode="group",
                xaxis_tickangle=-35,
                xaxis_title="Asignatura",
                yaxis_title="Horas totales",
                height=460,
                plot_bgcolor="#FFFFFF",
                paper_bgcolor="#F2F4F7",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                margin=dict(b=140, t=50, l=20, r=20),
                font=dict(size=10),
            )
            st.plotly_chart(fig_a2, use_container_width=True)

            st.divider()

            # ── Gráfico 3: mapa de calor ──────────────────────────────────────
            st.markdown("#### Mapa de Calor — Diferencia SHOA vs cada Currículo Internacional")

            delta_cols  = ["delta_padilla", "delta_sweden", "delta_uss", "delta_ucl"]
            intl_labels = ["Padilla (Colombia)", "Sweden", "USS", "UCL"]

            z_vals = df_asgn[delta_cols].values
            z_text = [[f"{v:+.1f}" for v in row] for row in z_vals]

            fig_a3 = go.Figure(go.Heatmap(
                z=z_vals,
                x=intl_labels,
                y=df_asgn["asignatura_shoa"],
                text=z_text,
                texttemplate="%{text}",
                colorscale=[
                    [0.0, "#2471A3"],
                    [0.5, "#FFFFFF"],
                    [1.0, "#C0392B"],
                ],
                zmid=0,
                colorbar=dict(title="Δ horas<br>(SHOA – Intl)"),
                hovertemplate="Asignatura: %{y}<br>%{x}: %{z:+.1f} h<extra></extra>",
            ))
            fig_a3.update_layout(
                xaxis_title="Currículo internacional",
                yaxis_title="Asignatura SHOA",
                height=max(420, 24 * len(df_asgn)),
                plot_bgcolor="#FFFFFF",
                paper_bgcolor="#F2F4F7",
                margin=dict(l=260, r=20, t=40, b=20),
                font=dict(size=10),
            )
            st.plotly_chart(fig_a3, use_container_width=True)

            st.divider()

            # ── Tabla interactiva ─────────────────────────────────────────────
            st.markdown("#### Tabla Completa de Asignaturas")

            col_f1, col_f2 = st.columns(2)
            with col_f1:
                fil_clf_a = st.multiselect(
                    "Filtrar por clasificación",
                    ["SOBREESTIMADA", "SUBESTIMADA", "ALINEADA", "EXCLUSIVA SHOA"],
                    default=["SOBREESTIMADA", "SUBESTIMADA", "ALINEADA", "EXCLUSIVA SHOA"],
                    key="fil_clf_a",
                )
            with col_f2:
                busqueda = st.text_input(
                    "Buscar asignatura", placeholder="Escribe para filtrar…", key="busq_a"
                )

            df_tbl_a = df_asgn[df_asgn["clasificacion"].isin(fil_clf_a)].copy()
            if busqueda:
                df_tbl_a = df_tbl_a[
                    df_tbl_a["asignatura_shoa"].str.contains(busqueda, case=False, na=False)
                ]

            def _color_clf_a(val):
                return {
                    "SOBREESTIMADA":  "background-color:#FADBD8;color:#C0392B",
                    "SUBESTIMADA":    "background-color:#D6EAF8;color:#2471A3",
                    "ALINEADA":       "background-color:#D5F5E3;color:#1A7A4A",
                    "EXCLUSIVA SHOA": "background-color:#E8DAEF;color:#6C3483",
                }.get(val, "")

            cols_show_a = [
                "asignatura_shoa", "horas_shoa", "intl_avg",
                "delta_h", "delta_pct",
                "horas_padilla", "horas_sweden", "horas_uss", "horas_ucl",
                "n_subtopicos", "clasificacion",
            ]
            df_show_a = df_tbl_a[cols_show_a].copy()
            df_show_a.columns = [
                "Asignatura SHOA", "Horas SHOA", "Prom. Intl",
                "Δ (h)", "Δ (%)",
                "Padilla", "Sweden", "USS", "UCL",
                "N° Tópicos", "Clasificación",
            ]

            try:
                styled_a = df_show_a.style.map(_color_clf_a, subset=["Clasificación"])
            except AttributeError:
                styled_a = df_show_a.style.applymap(_color_clf_a, subset=["Clasificación"])

            st.dataframe(
                styled_a.format({
                    "Horas SHOA": "{:.1f}", "Prom. Intl": "{:.1f}",
                    "Δ (h)": "{:+.1f}",    "Δ (%)": "{:+.1f}%",
                    "Padilla": "{:.1f}",   "Sweden": "{:.1f}",
                    "USS": "{:.1f}",       "UCL": "{:.1f}",
                }),
                use_container_width=True,
                height=430,
                hide_index=True,
            )
            st.markdown(f"**{len(df_tbl_a)}** asignaturas mostradas de {kpis_a['n_total']} totales")

            st.divider()

            # ── Exportación ───────────────────────────────────────────────────
            st.markdown("#### Exportar")
            xlsx_asgn = build_assignment_excel(df_asgn)
            st.download_button(
                label="📥 Exportar análisis por asignatura",
                data=xlsx_asgn,
                file_name=f"SHOA_Asignaturas_{date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

    # ── Footer sidebar ────────────────────────────────────────────────────
    st.sidebar.markdown("""
<hr style="border:none;border-top:1px solid rgba(200,168,75,0.5);margin-top:20px;">
<div style="text-align:center;padding:8px 0 4px;font-size:0.68rem;
            color:rgba(255,255,255,0.55);line-height:1.5;">
  © SHOA — Armada de Chile<br>
  Análisis IHO S-5A | Valparaíso
</div>
""", unsafe_allow_html=True)

    # ── Footer principal ──────────────────────────────────────────────────
    logo_ft = _logo_html(ancho=52)
    st.markdown(f"""
<div style="border-top:3px solid #C8A84B;padding:14px 0 6px;margin-top:8px;
            display:flex;align-items:center;justify-content:space-between;
            flex-wrap:wrap;gap:10px;">
  <div style="display:flex;align-items:center;gap:12px;">
    {logo_ft}
    <div>
      <p style="margin:0;font-size:0.78rem;color:#2C3E50;font-weight:600;">
        Servicio Hidrográfico y Oceanográfico de la Armada de Chile</p>
      <p style="margin:0;font-size:0.7rem;color:#888;">
        Errázuriz 254, Playa Ancha, Valparaíso &nbsp;·&nbsp; shoa.cl</p>
    </div>
  </div>
  <div style="text-align:right;">
    <p style="margin:0;font-size:0.7rem;color:#888;">
      Dashboard v1.0 &nbsp;·&nbsp; Análisis IHO S-5A<br>
      {len(df_analyzed)} sub-tópicos &nbsp;·&nbsp; Umbral ±15%</p>
  </div>
</div>
""", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
