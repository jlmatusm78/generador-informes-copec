from __future__ import annotations

from datetime import date, timedelta

import streamlit as st

from report_engine import ReportConfig, detect_default_week, generate_reports, load_data

st.set_page_config(
    page_title="Generador de Informes COPEC",
    page_icon="📊",
    layout="wide",
)

st.title("Generador Automático de Informes - Torre de Control COPEC")
st.caption("Carga el registro Guardian/FlotaGo y genera informes semanales en PDF, sin IRO.")

with st.sidebar:
    st.header("Configuración")
    uploaded = st.file_uploader("Archivo Excel", type=["xlsx", "xls"])
    comparison_weeks = st.selectbox("Semanas a comparar", [3, 4, 5, 6], index=1)
    include_global = st.checkbox("Informe Global COPEC", value=True)
    include_transportistas = st.checkbox("Informes por transportista", value=True)

if uploaded is None:
    st.info("Sube un archivo Excel con hojas GUARDIAN y/o FLOTAGO para comenzar.")
    st.stop()

try:
    data = load_data(uploaded)
except Exception as exc:
    st.error(f"No fue posible leer el archivo: {exc}")
    st.stop()

min_date = data["Fecha"].min().date()
max_date = data["Fecha"].max().date()
default_start, default_end = detect_default_week(data)

st.success(f"Archivo leído correctamente: {len(data):,} registros, desde {min_date:%d-%m-%Y} hasta {max_date:%d-%m-%Y}.")

col1, col2 = st.columns(2)
with col1:
    week_start = st.date_input("Inicio de semana", value=default_start, min_value=min_date, max_value=max_date)
with col2:
    week_end = st.date_input("Fin de semana", value=default_end, min_value=min_date, max_value=max_date)

if week_end < week_start:
    st.error("La fecha final no puede ser anterior a la fecha inicial.")
    st.stop()
if (week_end - week_start).days > 14:
    st.warning("El rango seleccionado supera 14 días. El formato está optimizado para informes semanales.")

current = data[(data["Fecha"].dt.date >= week_start) & (data["Fecha"].dt.date <= week_end)]

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Alertas", f"{len(current):,}")
m2.metric("Transportistas", current["Transportista"].nunique())
m3.metric("Conductores", current["Conductor"].nunique())
m4.metric("Equipos", current["Tracto"].nunique())
m5.metric("Fatiga", int((current["Tipo"] == "Fatiga").sum()))

st.subheader("Contenido de los informes")
st.write(
    "Evolución semanal, tendencias por tipo de alerta, semáforo ejecutivo, empresas y conductores críticos, "
    "fatiga, cumplimiento del protocolo, reincidencia, matriz empresa/tipo de alerta y plan de acción."
)

if st.button("Generar informes", type="primary", use_container_width=True):
    config = ReportConfig(
        week_start=week_start,
        week_end=week_end,
        comparison_weeks=comparison_weeks,
        include_global=include_global,
        include_transportistas=include_transportistas,
    )
    try:
        with st.spinner("Analizando datos y generando informes PDF..."):
            global_pdf, zip_bytes, summary = generate_reports(data, config)
    except Exception as exc:
        st.exception(exc)
        st.stop()

    st.success("Informes generados correctamente.")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Alertas", summary["total_alertas"])
    c2.metric("Transportistas", summary["transportistas"])
    c3.metric("Conductores", summary["conductores"])
    c4.metric("Equipos", summary["equipos"])
    c5.metric("Fatiga", summary["fatiga"])

    if global_pdf:
        st.download_button(
            "Descargar Informe Global COPEC",
            data=global_pdf,
            file_name=f"Informe_Global_COPEC_{week_start:%d%m}_{week_end:%d%m%Y}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    if zip_bytes:
        st.download_button(
            "Descargar ZIP completo",
            data=zip_bytes,
            file_name=f"Informes_COPEC_{week_start:%d%m}_{week_end:%d%m%Y}.zip",
            mime="application/zip",
            use_container_width=True,
        )
