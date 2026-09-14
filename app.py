from __future__ import annotations

from datetime import datetime
import io
import zipfile

from rco_sections import driver_count, rco_summary, fmt_pct

import pandas as pd
import streamlit as st

from email_engine import (
    SmtpSettings,
    build_dispatch_table,
    load_recipient_catalog,
    prepare_email,
    read_pdf_attachments,
    send_batch,
)
try:
    from report_engine import (
        ReportConfig,
        detect_default_month,
        detect_default_week,
        filter_real_transportistas,
        generate_reports,
        load_data,
    )
except ImportError as exc:
    raise ImportError(
        "No se pudo cargar la versión correcta de report_engine.py. "
        "Reemplaza en GitHub app.py y report_engine.py por los archivos del mismo paquete, "
        "elimina cualquier carpeta __pycache__ y reinicia la aplicación en Streamlit Cloud."
    ) from exc


st.set_page_config(
    page_title="Generador de Informes COPEC",
    page_icon="📊",
    layout="wide",
)

# Incrementar cuando cambia la estructura de los informes. Esto evita que una
# sesión abierta en Streamlit siga ofreciendo PDFs/ZIP generados con código
# anterior después de una actualización.
REPORT_SCHEMA_VERSION = "2026.09.14-sin-rco-v1"

if st.session_state.get("report_schema_version") != REPORT_SCHEMA_VERSION:
    for stale_key in (
        "global_pdf",
        "zip_bytes",
        "generation_summary",
        "dispatch_table",
        "send_log",
        "sent_keys",
    ):
        st.session_state.pop(stale_key, None)
    st.session_state.report_schema_version = REPORT_SCHEMA_VERSION

st.title("Generador Automático de Informes - Torre de Control COPEC")
st.caption("Genera informes operacionales semanales o mensuales con riesgo explicable, reincidencia, detalle de eventos y envío por correo.")

for key, default in {
    "global_pdf": None,
    "zip_bytes": None,
    "generation_summary": None,
    "dispatch_table": None,
    "send_log": [],
    "sent_keys": set(),
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

with st.sidebar:
    st.header("Configuración")
    uploaded = st.file_uploader("Archivo Excel de alertas", type=["xlsx", "xls"], key="alerts_file")
    report_mode = st.radio("Tipo de informe", ["Semanal", "Mensual"], horizontal=True)
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
default_start, default_end = detect_default_month(data) if report_mode == "Mensual" else detect_default_week(data)

with st.sidebar:
    if report_mode == "Semanal":
        comparison_periods = st.selectbox(
            "Semanas para comparar",
            [4, 5, 6],
            index=0,
            help="Cada semana comienza el lunes y termina el domingo.",
        )
    else:
        available_month_count = max(1, int(data["Fecha"].dt.to_period("M").nunique()))
        max_months = min(12, available_month_count)
        month_options = list(range(3, max_months + 1)) if max_months >= 3 else [max_months]
        default_months = min(6, max_months)
        comparison_periods = st.selectbox(
            "Meses para comparar",
            month_options,
            index=month_options.index(default_months) if default_months in month_options else 0,
            help="Se muestran hasta 12 meses, limitado por la información disponible.",
        )
    top_drivers = st.slider("Conductores en el ranking", 5, 30, 10, 1)

st.success(f"Archivo leído correctamente: {len(data):,} registros, desde {min_date:%d-%m-%Y} hasta {max_date:%d-%m-%Y}.")

if report_mode == "Mensual":
    available_months = sorted(data["Fecha"].dt.to_period("M").unique(), reverse=True)
    selected_month = st.selectbox(
        "Mes del informe",
        available_months,
        format_func=lambda p: p.strftime("%B %Y"),
        key="selected_report_month",
    )
    period_start = selected_month.start_time.date()
    period_end = selected_month.end_time.date()
    st.caption(f"El informe incluirá exclusivamente datos entre {period_start:%d-%m-%Y} y {period_end:%d-%m-%Y}.")
else:
    first_monday = min_date - pd.Timedelta(days=min_date.weekday())
    available_weeks = []
    cursor = default_start
    while cursor >= first_monday:
        available_weeks.append(cursor)
        cursor -= pd.Timedelta(days=7)
    period_start = st.selectbox(
        "Semana del informe",
        available_weeks,
        index=0,
        format_func=lambda start: f"Lun {start:%d-%m-%Y} → Dom {(start + pd.Timedelta(days=6)):%d-%m-%Y}",
    )
    period_end = period_start + pd.Timedelta(days=6)
    st.caption(f"Semana completa: lunes {period_start:%d-%m-%Y} a domingo {period_end:%d-%m-%Y}.")

if period_end < period_start:
    st.error("La fecha final no puede ser anterior a la fecha inicial.")
    st.stop()

report_data = filter_real_transportistas(data)
current = report_data[(report_data["Fecha"].dt.date >= period_start) & (report_data["Fecha"].dt.date <= period_end)]
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Alertas", f"{len(current):,}")
m2.metric("Transportistas", current["Transportista"].nunique())
m3.metric("Conductores identificados", driver_count(current))
m4.metric("Equipos", current["Tracto"].nunique())
m5.metric("Fatiga", int((current["Tipo"] == "Fatiga").sum()))

identification = rco_summary(current)
st.info(f"Sin RCO: {identification['sin_rco']} alertas ({fmt_pct(identification['pct'])}). "
        f"El análisis por conductor utiliza {identification['identified']} alertas identificadas.")
if identification['other']:
    st.warning(f"Otros registros sin identificación: {identification['other']}. Se muestran separados de Sin RCO.")

reports_tab, email_tab = st.tabs(["📄 Generación de informes", "✉️ Envío por correo"])

with reports_tab:
    st.subheader("Contenido de los informes")
    st.write(
        "Resumen ejecutivo, comparación histórica, riesgo operacional con cálculo explicado, reincidencia, ranking configurable, "
        "análisis técnico de Sensor tapado y Sensor desalineado, plan de acción y detalle completo de eventos. "
        "Cada transportista recibe un PDF completo con resumen, evolución, ranking y detalle."
    )

    if st.button("Generar informes", type="primary", use_container_width=True):
        config = ReportConfig(
            period_start=period_start,
            period_end=period_end,
            report_mode=report_mode,
            comparison_periods=comparison_periods,
            include_global=include_global,
            include_transportistas=include_transportistas,
            top_drivers=top_drivers,
        )
        try:
            with st.spinner("Analizando datos y generando informes PDF..."):
                global_pdf, zip_bytes, summary = generate_reports(data, config)
        except Exception as exc:
            st.exception(exc)
            st.stop()

        st.session_state.global_pdf = global_pdf
        st.session_state.zip_bytes = zip_bytes
        st.session_state.generation_summary = summary
        st.session_state.dispatch_table = None
        st.session_state.send_log = []
        st.session_state.sent_keys = set()
        st.success("Informes generados correctamente.")

    if st.session_state.generation_summary:
        summary = st.session_state.generation_summary
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Alertas", summary["total_alertas"])
        c2.metric("Transportistas", summary["transportistas"])
        c3.metric("Conductores identificados", summary["conductores"])
        c4.metric("Equipos", summary["equipos"])
        c5.metric("Riesgo / reincidencia", f"{summary['riesgo']} · {fmt_pct(summary['reincidencia'])}")

        if st.session_state.global_pdf:
            st.download_button(
                "Descargar Informe Global COPEC",
                data=st.session_state.global_pdf,
                file_name=f"Informe_Global_COPEC_{report_mode}_{period_start:%d%m}_{period_end:%d%m%Y}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        if st.session_state.zip_bytes:
            st.download_button(
                "Descargar ZIP completo",
                data=st.session_state.zip_bytes,
                file_name=f"Informes_COPEC_{report_mode}_{period_start:%d%m}_{period_end:%d%m%Y}.zip",
                mime="application/zip",
                use_container_width=True,
            )

with email_tab:
    st.subheader("Envío automático de informes")
    st.info(
        "Primero genera los informes. Luego carga la tabla de destinatarios, revisa la vista previa y envía una prueba antes del envío masivo."
    )

    if not st.session_state.zip_bytes:
        st.warning("Todavía no hay informes generados en esta sesión.")
        st.stop()

    recipients_file = st.file_uploader(
        "Tabla de destinatarios (Excel o CSV)",
        type=["xlsx", "xls", "csv"],
        help="Columnas: Transportista, Para, CC y Activo. Usa COPEC como transportista para el informe global.",
        key="recipients_file",
    )

    try:
        catalog = load_recipient_catalog(recipients_file) if recipients_file else pd.DataFrame(columns=["Transportista", "Para", "CC", "Activo"])
    except Exception as exc:
        st.error(f"No fue posible leer la tabla de destinatarios: {exc}")
        catalog = pd.DataFrame(columns=["Transportista", "Para", "CC", "Activo"])

    attachments = read_pdf_attachments(st.session_state.zip_bytes)
    if st.session_state.dispatch_table is None or recipients_file is not None:
        st.session_state.dispatch_table = build_dispatch_table(attachments, catalog)

    st.caption("Puedes corregir destinatarios directamente en la tabla antes de enviar.")
    edited = st.data_editor(
        st.session_state.dispatch_table,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Enviar": st.column_config.CheckboxColumn("Enviar"),
            "Transportista": st.column_config.TextColumn("Transportista", disabled=True),
            "Para": st.column_config.TextColumn("Para"),
            "CC": st.column_config.TextColumn("CC"),
            "Informe": st.column_config.TextColumn("Informe", disabled=True),
            "Estado": st.column_config.TextColumn("Estado", disabled=True),
        },
        key="dispatch_editor",
    )
    edited["Estado"] = edited["Para"].apply(lambda x: "Listo" if str(x).strip() else "Falta destinatario")
    st.session_state.dispatch_table = edited

    selected = edited[edited["Enviar"]].copy()
    missing = selected[selected["Para"].astype(str).str.strip() == ""]
    c1, c2, c3 = st.columns(3)
    c1.metric("Informes disponibles", len(attachments))
    c2.metric("Seleccionados", len(selected))
    c3.metric("Sin destinatario", len(missing))

    if len(selected):
        preview_name = st.selectbox("Vista previa", selected["Transportista"].tolist())
        preview_row = selected[selected["Transportista"] == preview_name].iloc[0]
        st.markdown(f"**Para:** {preview_row['Para'] or 'Sin destinatario'}")
        st.markdown(f"**CC:** {preview_row['CC'] or 'Sin copia'}")
        st.markdown(f"**Adjunto:** {preview_row['Informe']}")
        st.markdown(
            f"**Asunto:** Informe {report_mode.lower()} de alertas | {preview_name} | "
            f"{period_start:%d-%m-%Y} al {period_end:%d-%m-%Y}"
        )

    st.divider()
    st.subheader("Configuración de Gmail / Google Workspace")
    try:
        gmail_secrets = st.secrets.get("gmail", {})
    except Exception:
        gmail_secrets = {}

    configured_sender = str(gmail_secrets.get("sender_email", ""))
    if configured_sender:
        st.success(f"Cuenta configurada: {configured_sender}")
    else:
        st.warning("Falta configurar la sección [gmail] en Streamlit Secrets. Consulta el README incluido.")

    test_recipient = st.text_input("Correo para envío de prueba", placeholder="tu-correo@empresa.cl")
    confirm = st.checkbox("He revisado los destinatarios, los adjuntos y el período de los informes.")

    attachment_lookup = {item["filename"]: item for item in attachments}

    def smtp_settings_from_secrets() -> SmtpSettings:
        if not gmail_secrets:
            raise ValueError("No se encontró la configuración [gmail] en Streamlit Secrets.")
        required = ["username", "password", "sender_email"]
        missing_keys = [k for k in required if not str(gmail_secrets.get(k, "")).strip()]
        if missing_keys:
            raise ValueError("Faltan secretos Gmail: " + ", ".join(missing_keys))
        return SmtpSettings(
            host=str(gmail_secrets.get("host", "smtp.gmail.com")),
            port=int(gmail_secrets.get("port", 587)),
            username=str(gmail_secrets["username"]),
            password=str(gmail_secrets["password"]),
            sender_email=str(gmail_secrets["sender_email"]),
            sender_name=str(gmail_secrets.get("sender_name", "Torre de Control COPEC")),
            use_tls=bool(gmail_secrets.get("use_tls", True)),
            use_ssl=bool(gmail_secrets.get("use_ssl", False)),
        )

    b1, b2 = st.columns(2)
    with b1:
        if st.button("Enviar correo de prueba", use_container_width=True, disabled=not (confirm and test_recipient and len(selected))):
            try:
                settings = smtp_settings_from_secrets()
                row = selected.iloc[0].to_dict()
                attachment = attachment_lookup[row["Informe"]]
                email = prepare_email(
                    row,
                    attachment,
                    period_start.strftime("%d-%m-%Y"),
                    period_end.strftime("%d-%m-%Y"),
                    test_recipient=test_recipient,
                )
                results = send_batch(settings, [email])
                st.session_state.send_log.extend(results)
                if results[0]["Resultado"] == "Enviado":
                    st.success("Correo de prueba enviado correctamente.")
                else:
                    st.error(results[0]["Detalle"])
            except Exception as exc:
                st.error(str(exc))

    with b2:
        if st.button("Enviar informes seleccionados", type="primary", use_container_width=True, disabled=not (confirm and len(selected) and not len(missing))):
            try:
                settings = smtp_settings_from_secrets()
                prepared = []
                skipped = 0
                for row in selected.to_dict("records"):
                    key = f"{report_mode}|{period_start}|{period_end}|{row['Informe']}|{row['Para']}"
                    if key in st.session_state.sent_keys:
                        skipped += 1
                        continue
                    prepared.append(prepare_email(
                        row,
                        attachment_lookup[row["Informe"]],
                        period_start.strftime("%d-%m-%Y"),
                        period_end.strftime("%d-%m-%Y"),
                    ))
                if not prepared:
                    st.warning("Todos los informes seleccionados ya fueron enviados durante esta sesión.")
                else:
                    with st.spinner(f"Enviando {len(prepared)} correos..."):
                        results = send_batch(settings, prepared)
                    st.session_state.send_log.extend(results)
                    for result, email in zip(results, prepared):
                        if result["Resultado"] == "Enviado":
                            key = f"{report_mode}|{period_start}|{period_end}|{email.attachment_name}|{';'.join(email.to)}"
                            st.session_state.sent_keys.add(key)
                    ok = sum(r["Resultado"] == "Enviado" for r in results)
                    errors = len(results) - ok
                    st.success(f"Enviados: {ok}. Errores: {errors}. Omitidos por duplicado: {skipped}.")
            except Exception as exc:
                st.error(str(exc))

    if st.session_state.send_log:
        st.subheader("Registro de envíos")
        log_df = pd.DataFrame(st.session_state.send_log)
        st.dataframe(log_df, use_container_width=True, hide_index=True)
        st.download_button(
            "Descargar registro CSV",
            data=log_df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"Registro_envios_{datetime.now():%Y%m%d_%H%M}.csv",
            mime="text/csv",
            use_container_width=True,
        )