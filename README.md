# Dashboard de Análisis Curricular — SHOA

Aplicación web interactiva que compara el currículo de hidrografía del **SHOA (Chile)** con cuatro currículos internacionales de referencia: Almirante Padilla (Colombia), Sweden, USS y University College London (UCL).

🔗 **Dashboard en línea:** [Ver dashboard](https://jsantistavilo-maker-shoa-curriculum-analysis.streamlit.app)

---

## Funcionalidades

| Tab | Contenido |
|-----|-----------|
| 📊 Vista General / KPIs | Horas totales, índice de convergencia, top módulos críticos |
| 📈 Comparativa de Horas | Gráfico de barras interactivo, tabla comparativa con código de colores |
| 🔥 Análisis de Diferencias | Heatmap de deltas, barras horizontales, radar chart por tópico |
| ⚠️ Módulos Críticos | Lista priorizada de módulos sobrevalorados y subvalorados |
| 💡 Recomendaciones | Recomendaciones de ajuste + exportación a Excel |
| 📚 Análisis por Asignatura | Análisis agrupado por nombre de asignatura con comparativa |

## Lógica de clasificación (módulos)

- **SOBREVALORADO**: SHOA tiene >15% más horas que el promedio internacional
- **SUBVALORADO**: SHOA tiene >15% menos horas que el promedio internacional
- **ALINEADO**: diferencia dentro del ±15%

## Lógica de clasificación (asignaturas)

- **SOBREESTIMADA**: SHOA tiene >20% más horas que el promedio internacional
- **SUBESTIMADA**: SHOA tiene >20% menos horas
- **ALINEADA**: diferencia dentro del ±20%
- **EXCLUSIVA SHOA**: no tiene equivalente en los currículos internacionales

---

## Para desarrolladores — ejecución local

### Requisitos
- Python 3.11 o superior

### Instalación

```powershell
# 1. Clonar el repositorio
git clone https://github.com/jsantistavilo-maker/shoa-curriculum-analysis.git
cd shoa-curriculum-analysis

# 2. Crear entorno virtual
python -m venv venv
.\venv\Scripts\Activate.ps1

# 3. Instalar dependencias
pip install -r requirements.txt
```

### Generar datos desde el Excel original

Coloca el archivo `Tabla resumen.xlsx` en:
```
~/OneDrive/Desktop/shoa resumen/Tabla resumen.xlsx
```

Luego ejecuta:
```powershell
python generate_data.py
```

Esto genera `data/curriculum_data.json` con todos los datos procesados.

### Ejecutar la app

```powershell
streamlit run app.py
```

### Actualizar datos (cuando cambia el Excel)

```powershell
python generate_data.py
git add data/curriculum_data.json
git commit -m "Actualizar datos curriculares"
git push
```

Streamlit Cloud detecta el push y actualiza el dashboard automáticamente.

---

## Estructura del proyecto

```
shoa-curriculum-analysis/
├── app.py                  ← Aplicación Streamlit principal (6 tabs)
├── data_loader.py          ← Carga desde JSON (online) o Excel (local)
├── analysis.py             ← Lógica de comparación y clasificación
├── recommendations.py      ← Recomendaciones y exportación Excel
├── priority_sheet.py       ← Hoja de asignaturas prioritarias
├── assignment_analysis.py  ← Análisis por asignatura
├── validate_data.py        ← Validación de datos (solo modo local)
├── requirements.txt
├── README.md
├── .gitignore
├── .streamlit/
│   └── config.toml
└── data/
    └── curriculum_data.json   ← Datos procesados (generado por generate_data.py)
```

> `generate_data.py` y `*.xlsx` están en `.gitignore` y no se suben al repositorio.
