from __future__ import annotations

import io
import os
import re
import tempfile
import textwrap
import unicodedata
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Iterable

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

ALERT_WEIGHTS = {
    "Fatiga": 5,
    "Uso celular": 4,
    "Sin cinturón": 3,
    "Tapado cámara": 3,
    "Conductor fumando": 2,
    "Sin conductor": 2,
    "Cámara desalineada": 1,
    "Bostezo": 1,
}

GENERIC_TRANSPORTISTAS = [
    "COPEC", "NO ES COPEC", "PLANTA", "DESCONOCIDO", "OWL", "GPS",
    "PRUEBA", "SIN TRANSPORTISTA",
]

@dataclass
class ReportConfig:
    week_start: date
    week_end: date
    comparison_weeks: int = 4
    include_global: bool = True
    include_transportistas: bool = True


def norm_text(value: object) -> str:
    text = str(value).upper().strip()
    text = "".join(
        c for c in unicodedata.normalize("NFD", text)
        if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"\s+", " ", text)


def safe_filename(value: object, max_len: int = 90) -> str:
    text = "".join(
        c for c in unicodedata.normalize("NFD", str(value))
        if unicodedata.category(c) != "Mn"
    )
    text = re.sub(r"[^\w\s\-.&]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return (text[:max_len] or "SIN_NOMBRE").strip()


def truncate(value: object, length: int) -> str:
    text = str(value)
    return text if len(text) <= length else text[: length - 1] + "…"


def wrap_label(value: object, width: int = 20) -> str:
    return "\n".join(textwrap.wrap(str(value), width=width))


def normalize_alert(value: object) -> str:
    s = str(value).lower()
    if any(x in s for x in ["cansancio", "fatiga", "somnolencia"]):
        return "Fatiga"
    if any(x in s for x in ["celular", "telefono", "teléfono"]):
        return "Uso celular"
    if "cintur" in s:
        return "Sin cinturón"
    if "desalinead" in s:
        return "Cámara desalineada"
    if "tapad" in s:
        return "Tapado cámara"
    if "bostezo" in s:
        return "Bostezo"
    if "fumando" in s:
        return "Conductor fumando"
    if "sin conductor" in s:
        return "Sin conductor"
    return str(value).strip()[:55] or "Sin clasificar"


def load_data(uploaded_file) -> pd.DataFrame:
    raw = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else Path(uploaded_file).read_bytes()
    excel = pd.ExcelFile(io.BytesIO(raw))
    frames: list[pd.DataFrame] = []
    for sheet in excel.sheet_names:
        name = sheet.strip().upper()
        if name in {"GUARDIAN", "FLOTAGO"}:
            df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet)
            df["Plataforma"] = "FLOTAGO" if "FLOTA" in name else "GUARDIAN"
            frames.append(df)
    if not frames:
        raise ValueError("El archivo debe contener hojas llamadas GUARDIAN y/o FLOTAGO.")

    data = pd.concat(frames, ignore_index=True)
    required = [
        "ID", "Fecha", "Transportista", "Conductor", "Tracto", "Incidente",
        "Plataforma", "Conductor se detine mínimo 15 minutos",
    ]
    for col in required:
        if col not in data.columns:
            data[col] = np.nan

    data["Fecha"] = pd.to_datetime(data["Fecha"], errors="coerce")
    data = data.dropna(subset=["Fecha"]).copy()
    data["Transportista"] = data["Transportista"].fillna("SIN TRANSPORTISTA").astype(str).str.strip()
    data["Conductor"] = data["Conductor"].fillna("SIN CONDUCTOR").astype(str).str.strip()
    data["Tracto"] = data["Tracto"].fillna("S/I").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    data["Incidente"] = data["Incidente"].fillna("Sin clasificar").astype(str).str.strip()
    data["Plataforma"] = data["Plataforma"].fillna("SIN PLATAFORMA").astype(str).str.upper().str.strip()
    data["Tipo"] = data["Incidente"].apply(normalize_alert)
    data["FechaDia"] = data["Fecha"].dt.date
    return data.sort_values("Fecha").reset_index(drop=True)


def filter_real_transportistas(data: pd.DataFrame) -> pd.DataFrame:
    mask = data["Transportista"].map(
        lambda x: not any(g in norm_text(x) for g in GENERIC_TRANSPORTISTAS)
    )
    return data[mask].copy()


def detect_default_week(data: pd.DataFrame) -> tuple[date, date]:
    max_date = data["Fecha"].max().date()
    # Last complete Monday-Sunday week ending on or before max_date.
    end = max_date - timedelta(days=(max_date.weekday() + 1) % 7)
    start = end - timedelta(days=6)
    return start, end


def comparison_ranges(start: date, count: int) -> list[tuple[str, pd.Timestamp, pd.Timestamp]]:
    ranges = []
    first = start - timedelta(days=7 * (count - 1))
    for i in range(count):
        s = first + timedelta(days=7 * i)
        e = s + timedelta(days=6)
        label = f"{s.strftime('%d-%m')} / {e.strftime('%d-%m')}"
        ranges.append((label, pd.Timestamp(s), pd.Timestamp(datetime.combine(e, datetime.max.time()))))
    return ranges


def pct_change(current: float, previous: float) -> str:
    if previous == 0 and current == 0:
        return "0%"
    if previous == 0:
        return "+100%"
    return f"{((current - previous) / previous) * 100:+.1f}%"


def compliance_summary(df: pd.DataFrame) -> tuple[int, int, int, int, float]:
    fatigue = df[df["Tipo"].eq("Fatiga")]
    total = len(fatigue)
    if total == 0:
        return 0, 0, 0, 0, 0.0
    values = (
        fatigue["Conductor se detine mínimo 15 minutos"]
        .fillna("PENDIENTE").astype(str).str.upper().str.strip()
    )
    yes = int((values == "SI").sum())
    no = int((values == "NO").sum())
    pending = total - yes - no
    return total, yes, no, pending, yes / total * 100


def driver_ranking(df: pd.DataFrame, topn: int = 10, include_transportista: bool = False) -> pd.DataFrame:
    if df.empty:
        columns = ["Conductor", "Alertas", "Dias", "Fatiga", "Celular", "Tendencia"]
        if include_transportista:
            columns.insert(1, "Transportista")
        return pd.DataFrame(columns=columns)
    group = ["Conductor", "Transportista"] if include_transportista else ["Conductor"]
    base = df.groupby(group).agg(
        Alertas=("ID", "count"),
        Dias=("FechaDia", "nunique"),
        Fatiga=("Tipo", lambda s: int((s == "Fatiga").sum())),
        Celular=("Tipo", lambda s: int((s == "Uso celular").sum())),
    ).reset_index()
    return base.sort_values(["Alertas", "Fatiga", "Celular", "Dias"], ascending=False).head(topn)


def recurrence_summary(df: pd.DataFrame) -> dict[str, float]:
    counts = df.groupby("Conductor").size() if not df.empty else pd.Series(dtype=int)
    one = int((counts == 1).sum())
    two = int((counts == 2).sum())
    three_plus = int((counts >= 3).sum())
    total = int(len(counts))
    rate = (three_plus / total * 100) if total else 0.0
    return {"one": one, "two": two, "three_plus": three_plus, "total": total, "rate": rate}


def _styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="BodyCustom", parent=styles["BodyText"], fontSize=8.2, leading=10.8))
    styles.add(ParagraphStyle(name="SmallCustom", parent=styles["BodyText"], fontSize=6.2, leading=7.7))
    styles.add(ParagraphStyle(name="TitleCustom", parent=styles["Title"], fontSize=17, leading=20, textColor=colors.HexColor("#1F4E79")))
    styles.add(ParagraphStyle(name="H2Custom", parent=styles["Heading2"], fontSize=11, leading=14, textColor=colors.HexColor("#1F4E79")))
    return styles


def _style_table(table: Table, font_size: float = 6.2) -> Table:
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#C9D6E2")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2.5),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return table


def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.grey)
    canvas.drawString(1.4 * cm, 0.8 * cm, "Torre de Control COPEC - Informe automático")
    canvas.drawRightString(A4[0] - 1.4 * cm, 0.8 * cm, f"Página {doc.page}")
    canvas.restoreState()


def _plot_line(labels: list[str], series: dict[str, list[int]], title: str, ylabel: str, path: Path) -> None:
    plt.figure(figsize=(8.6, 4.0))
    for label, values in series.items():
        plt.plot(labels, values, marker="o", label=label)
    plt.title(title)
    plt.ylabel(ylabel)
    plt.xticks(rotation=15)
    plt.grid(True, alpha=0.25)
    if len(series) > 1:
        plt.legend(fontsize=7.5, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.2))
    plt.tight_layout()
    plt.savefig(path, dpi=140, bbox_inches="tight")
    plt.close()


def _plot_bar(labels: list[str], values: Iterable[int], title: str, xlabel: str, path: Path) -> None:
    plt.figure(figsize=(8.2, 4.8))
    plt.barh(labels[::-1], list(values)[::-1])
    plt.title(title)
    plt.xlabel(xlabel)
    plt.tight_layout()
    plt.savefig(path, dpi=140)
    plt.close()


def _plot_heatmap(matrix: pd.DataFrame, title: str, path: Path) -> None:
    if matrix.empty:
        matrix = pd.DataFrame([[0]], index=["Sin datos"], columns=["Sin datos"])
    plt.figure(figsize=(9.2, 5.4))
    plt.imshow(matrix.values, aspect="auto")
    plt.xticks(range(len(matrix.columns)), [wrap_label(c, 13) for c in matrix.columns], rotation=30, ha="right", fontsize=7)
    plt.yticks(range(len(matrix.index)), [truncate(x, 30) for x in matrix.index], fontsize=7)
    for i in range(matrix.shape[0]):
        for j in range(matrix.shape[1]):
            value = int(matrix.values[i, j])
            if value:
                plt.text(j, i, str(value), ha="center", va="center", fontsize=7)
    plt.title(title)
    plt.tight_layout()
    plt.savefig(path, dpi=140)
    plt.close()


def _week_slice(data: pd.DataFrame, start: date, end: date) -> pd.DataFrame:
    start_ts = pd.Timestamp(start)
    end_ts = pd.Timestamp(datetime.combine(end, datetime.max.time()))
    return data[(data["Fecha"] >= start_ts) & (data["Fecha"] <= end_ts)].copy()


def _weekly_metrics(data: pd.DataFrame, ranges) -> tuple[list[int], dict[str, list[int]]]:
    totals = []
    types = data["Tipo"].value_counts().head(5).index.tolist()
    by_type = {t: [] for t in types}
    for _, start, end in ranges:
        week = data[(data["Fecha"] >= start) & (data["Fecha"] <= end)]
        totals.append(len(week))
        for t in types:
            by_type[t].append(int((week["Tipo"] == t).sum()))
    return totals, by_type


def generate_global_report(data: pd.DataFrame, config: ReportConfig, output_dir: Path) -> Path:
    styles = _styles()
    current = _week_slice(data, config.week_start, config.week_end)
    ranges = comparison_ranges(config.week_start, config.comparison_weeks)
    labels = [r[0] for r in ranges]
    totals, by_type = _weekly_metrics(data, ranges)
    previous = data[(data["Fecha"] >= ranges[-2][1]) & (data["Fecha"] <= ranges[-2][2])] if len(ranges) > 1 else current.iloc[0:0]

    charts = output_dir / "charts_global"
    charts.mkdir(parents=True, exist_ok=True)
    p_evol = charts / "evolucion_total.png"
    p_types = charts / "evolucion_tipos.png"
    p_comp = charts / "empresas.png"
    p_drv = charts / "conductores.png"
    p_heat = charts / "heatmap.png"

    _plot_line(labels, {"Alertas": totals}, "Evolución semanal de alertas", "Alertas", p_evol)
    _plot_line(labels, by_type or {"Sin datos": [0] * len(labels)}, "Evolución semanal por tipo de alerta", "Alertas", p_types)

    rank_trans = current.groupby("Transportista").agg(
        Alertas=("ID", "count"), Conductores=("Conductor", "nunique"),
        Equipos=("Tracto", "nunique"), Fatiga=("Tipo", lambda s: int((s == "Fatiga").sum())),
    ).reset_index().sort_values(["Alertas", "Fatiga"], ascending=False)
    prev_trans = previous.groupby("Transportista").size()
    rank_trans["Semana anterior"] = rank_trans["Transportista"].map(prev_trans).fillna(0).astype(int)
    rank_trans["Tendencia"] = np.where(rank_trans["Alertas"] > rank_trans["Semana anterior"], "↑", np.where(rank_trans["Alertas"] < rank_trans["Semana anterior"], "↓", "↔"))

    drivers = driver_ranking(current, 12, include_transportista=True)
    prev_drv = previous.groupby("Conductor").size()
    drivers["Semana anterior"] = drivers["Conductor"].map(prev_drv).fillna(0).astype(int)
    drivers["Tendencia"] = np.where(drivers["Alertas"] > drivers["Semana anterior"], "↑", np.where(drivers["Alertas"] < drivers["Semana anterior"], "↓", "↔"))

    _plot_bar([wrap_label(x, 24) for x in rank_trans.head(10)["Transportista"]], rank_trans.head(10)["Alertas"], "Top empresas por alertas", "Alertas", p_comp)
    _plot_bar([wrap_label(f"{r.Conductor} ({r.Transportista})", 30) for r in drivers.head(10).itertuples()], drivers.head(10)["Alertas"], "Top conductores por alertas", "Alertas", p_drv)

    top_companies = rank_trans.head(10)["Transportista"].tolist()
    top_types = current["Tipo"].value_counts().head(6).index.tolist()
    matrix = current[current["Transportista"].isin(top_companies) & current["Tipo"].isin(top_types)].pivot_table(
        index="Transportista", columns="Tipo", values="ID", aggfunc="count", fill_value=0
    )
    if top_companies and top_types:
        matrix = matrix.reindex(index=top_companies, columns=top_types).fillna(0)
    _plot_heatmap(matrix, "Matriz empresa vs tipo de alerta", p_heat)

    total = len(current)
    guardian = int((current["Plataforma"] == "GUARDIAN").sum())
    flotago = int((current["Plataforma"] == "FLOTAGO").sum())
    fatigue_total, fatigue_yes, fatigue_no, fatigue_pending, fatigue_pct = compliance_summary(current)
    recurrence = recurrence_summary(current)
    previous_total = len(previous)
    variation = pct_change(total, previous_total)

    pdf = output_dir / f"00_Informe_Global_COPEC_{config.week_start:%d%m}_{config.week_end:%d%m%Y}.pdf"
    doc = SimpleDocTemplate(str(pdf), pagesize=A4, rightMargin=1.1 * cm, leftMargin=1.1 * cm, topMargin=1.0 * cm, bottomMargin=1.2 * cm)
    elements = []

    elements += [
        Paragraph("INFORME SEMANAL GLOBAL COPEC", styles["TitleCustom"]),
        Paragraph(f"Período: {config.week_start:%d-%m-%Y} al {config.week_end:%d-%m-%Y} - Guardian y FlotaGO", styles["BodyCustom"]),
        Spacer(1, 8), Paragraph("1. Dashboard ejecutivo", styles["H2Custom"]),
    ]
    kpi = [
        ["Indicador", "Valor"], ["Total alertas", total], ["Variación semanal", variation],
        ["Guardian", guardian], ["FlotaGO", flotago], ["Transportistas", current["Transportista"].nunique()],
        ["Conductores", current["Conductor"].nunique()], ["Equipos", current["Tracto"].nunique()],
        ["Fatiga", fatigue_total], ["Cumplimiento fatiga", f"{fatigue_pct:.1f}%"],
        ["Conductores reincidentes (3+)", recurrence["three_plus"]], ["% reincidencia", f"{recurrence['rate']:.1f}%"],
    ]
    sem = [
        ["Indicador", "Actual", "Semana anterior", "Tendencia"],
        ["Alertas", total, previous_total, "↓" if total < previous_total else ("↑" if total > previous_total else "↔")],
        ["Fatiga", fatigue_total, int((previous["Tipo"] == "Fatiga").sum()), "↓" if fatigue_total < int((previous["Tipo"] == "Fatiga").sum()) else "↑"],
        ["Reincidentes", recurrence["three_plus"], recurrence_summary(previous)["three_plus"], "↓" if recurrence["three_plus"] < recurrence_summary(previous)["three_plus"] else "↑"],
    ]
    elements += [
        Table([[_style_table(Table(kpi, colWidths=[4.2 * cm, 3.0 * cm]), 6.5), _style_table(Table(sem, colWidths=[3.5 * cm, 2.0 * cm, 2.4 * cm, 1.3 * cm]), 6.2)]], colWidths=[7.5 * cm, 10.4 * cm], style=[("VALIGN", (0, 0), (-1, -1), "TOP")]),
        Spacer(1, 8), Image(str(p_evol), width=18 * cm, height=7.2 * cm), PageBreak(),
    ]

    elements += [
        Paragraph("TENDENCIAS OPERACIONALES", styles["TitleCustom"]),
        Paragraph("2. Evolución semanal por tipo de alerta", styles["H2Custom"]),
        Image(str(p_types), width=18 * cm, height=8.2 * cm), Spacer(1, 8),
        Paragraph(f"La operación registra una variación de {variation} respecto de la semana anterior. La lectura debe enfocarse en los tipos de alerta que presentan tendencia al alza.", styles["BodyCustom"]),
        PageBreak(),
    ]

    comp_rows = [["Empresa", "Alertas", "% total", "Sem. ant.", "Tend.", "Conductores", "Fatiga"]]
    for _, r in rank_trans.head(10).iterrows():
        comp_rows.append([
            Paragraph(truncate(r["Transportista"], 35), styles["SmallCustom"]), int(r["Alertas"]),
            f"{(r['Alertas'] / max(total, 1) * 100):.1f}%", int(r["Semana anterior"]), r["Tendencia"],
            int(r["Conductores"]), int(r["Fatiga"]),
        ])
    elements += [
        Paragraph("EMPRESAS DE TRANSPORTE", styles["TitleCustom"]),
        Paragraph("3. Empresas con mayor cantidad de alertas", styles["H2Custom"]),
        Image(str(p_comp), width=18 * cm, height=8.0 * cm), Spacer(1, 8),
        _style_table(Table(comp_rows, colWidths=[5.8 * cm, 1.4 * cm, 1.5 * cm, 1.6 * cm, 1.1 * cm, 1.8 * cm, 1.3 * cm]), 5.7),
        PageBreak(),
    ]

    drv_rows = [["Conductor", "Empresa", "Alertas", "Días", "Fatiga", "Celular", "Sem. ant.", "Tend."]]
    for _, r in drivers.head(10).iterrows():
        drv_rows.append([
            Paragraph(truncate(r["Conductor"], 28), styles["SmallCustom"]),
            Paragraph(truncate(r["Transportista"], 28), styles["SmallCustom"]),
            int(r["Alertas"]), int(r["Dias"]), int(r["Fatiga"]), int(r["Celular"]), int(r["Semana anterior"]), r["Tendencia"],
        ])
    elements += [
        Paragraph("CONDUCTORES CRÍTICOS", styles["TitleCustom"]),
        Paragraph("4. Conductores con mayor concentración de alertas", styles["H2Custom"]),
        Image(str(p_drv), width=18 * cm, height=8.0 * cm), Spacer(1, 8),
        _style_table(Table(drv_rows, colWidths=[4.2 * cm, 4.0 * cm, 1.3 * cm, 1.0 * cm, 1.2 * cm, 1.2 * cm, 1.5 * cm, 1.1 * cm]), 5.5),
        PageBreak(),
    ]

    fatigue_rows = [
        ["Indicador", "Resultado"], ["Alertas de fatiga", fatigue_total], ["Cumplen 15 minutos", fatigue_yes],
        ["No cumplen", fatigue_no], ["Pendientes", fatigue_pending], ["Cumplimiento", f"{fatigue_pct:.1f}%"],
    ]
    recurrence_rows = [
        ["Indicador", "Conductores"], ["Con 1 alerta", recurrence["one"]], ["Con 2 alertas", recurrence["two"]],
        ["Con 3 o más alertas", recurrence["three_plus"]], ["% reincidencia", f"{recurrence['rate']:.1f}%"],
    ]
    elements += [
        Paragraph("GESTIÓN DE FATIGA Y REINCIDENCIA", styles["TitleCustom"]),
        Paragraph("5. Cumplimiento del protocolo y concentración de alertas", styles["H2Custom"]),
        Table([[_style_table(Table(fatigue_rows, colWidths=[5.8 * cm, 3.0 * cm]), 6.4), _style_table(Table(recurrence_rows, colWidths=[5.2 * cm, 3.2 * cm]), 6.4)]], colWidths=[9.2 * cm, 8.7 * cm], style=[("VALIGN", (0, 0), (-1, -1), "TOP")]),
        Spacer(1, 10), Paragraph("La reincidencia identifica si las alertas se concentran en un grupo reducido de conductores. Se consideran reincidentes quienes registran tres o más eventos en la semana.", styles["BodyCustom"]),
        PageBreak(),
    ]

    elements += [
        Paragraph("HALLAZGOS OPERACIONALES", styles["TitleCustom"]),
        Paragraph("6. Matriz empresa y tipo de alerta", styles["H2Custom"]),
        Image(str(p_heat), width=18 * cm, height=9.6 * cm), PageBreak(),
    ]

    actions = [
        ["Hallazgo", "Acción recomendada"],
        ["Empresas con aumento de alertas", "Solicitar análisis causal y plan correctivo semanal."],
        ["Conductores reincidentes", "Realizar retroalimentación individual y seguimiento de reincidencia."],
        ["Aumento de fatiga", "Reforzar gestión inmediata y cumplimiento del descanso mínimo."],
        ["Alertas de cámara", "Revisar instalación, mantención y posible manipulación del dispositivo."],
        ["Uso de celular / cinturón", "Aplicar campañas focalizadas y medidas correctivas con el transportista."],
    ]
    elements += [
        Paragraph("CONCLUSIONES Y PLAN DE ACCIÓN", styles["TitleCustom"]),
        Paragraph("7. Priorización para la semana siguiente", styles["H2Custom"]),
        _style_table(Table([[Paragraph(str(a), styles["SmallCustom"]), Paragraph(str(b), styles["SmallCustom"])] for a, b in actions], colWidths=[5.3 * cm, 11.7 * cm]), 6.0),
        Spacer(1, 10), Paragraph(f"La semana analizada registra {total} alertas y una variación de {variation}. La prioridad debe centrarse en empresas con tendencia al alza, conductores reincidentes y eventos de fatiga sin cumplimiento del protocolo.", styles["BodyCustom"]),
    ]
    doc.build(elements, onFirstPage=_footer, onLaterPages=_footer)
    return pdf


def generate_transportista_report(data: pd.DataFrame, transportista: str, config: ReportConfig, output_dir: Path, index: int) -> Path:
    styles = _styles()
    all_t = data[data["Transportista"] == transportista].copy()
    current = _week_slice(all_t, config.week_start, config.week_end)
    ranges = comparison_ranges(config.week_start, config.comparison_weeks)
    labels = [r[0] for r in ranges]
    totals, by_type = _weekly_metrics(all_t, ranges)
    previous = all_t[(all_t["Fecha"] >= ranges[-2][1]) & (all_t["Fecha"] <= ranges[-2][2])] if len(ranges) > 1 else current.iloc[0:0]

    chart_dir = output_dir / f"charts_{index:02d}"
    chart_dir.mkdir(parents=True, exist_ok=True)
    p_evol = chart_dir / "evolucion.png"
    p_types = chart_dir / "tipos.png"
    p_drivers = chart_dir / "conductores.png"
    _plot_line(labels, {"Alertas": totals}, "Evolución semanal de alertas", "Alertas", p_evol)
    _plot_line(labels, by_type or {"Sin datos": [0] * len(labels)}, "Evolución semanal por tipo de alerta", "Alertas", p_types)

    drivers = driver_ranking(current, 10, include_transportista=False)
    prev_drv = previous.groupby("Conductor").size()
    drivers["Semana anterior"] = drivers["Conductor"].map(prev_drv).fillna(0).astype(int)
    drivers["Tendencia"] = np.where(drivers["Alertas"] > drivers["Semana anterior"], "↑", np.where(drivers["Alertas"] < drivers["Semana anterior"], "↓", "↔"))
    _plot_bar([wrap_label(x, 25) for x in drivers.head(10)["Conductor"]], drivers.head(10)["Alertas"], "Conductores con más alertas", "Alertas", p_drivers)

    total = len(current)
    previous_total = len(previous)
    fatigue_total, fatigue_yes, fatigue_no, fatigue_pending, fatigue_pct = compliance_summary(current)
    recurrence = recurrence_summary(current)
    variation = pct_change(total, previous_total)

    pdf = output_dir / f"{index:02d}_Informe_{safe_filename(transportista)}_{config.week_start:%d%m}_{config.week_end:%d%m%Y}.pdf"
    doc = SimpleDocTemplate(str(pdf), pagesize=A4, rightMargin=1.1 * cm, leftMargin=1.1 * cm, topMargin=1.0 * cm, bottomMargin=1.2 * cm)
    elements = [
        Paragraph("INFORME SEMANAL DE ALERTAS", styles["TitleCustom"]),
        Paragraph(f"Transportista: {transportista}<br/>Período: {config.week_start:%d-%m-%Y} al {config.week_end:%d-%m-%Y}", styles["BodyCustom"]),
        Spacer(1, 8), Paragraph("1. Resumen ejecutivo", styles["H2Custom"]),
    ]
    kpi = [
        ["Indicador", "Valor"], ["Total alertas", total], ["Semana anterior", previous_total], ["Variación", variation],
        ["Guardian", int((current["Plataforma"] == "GUARDIAN").sum())], ["FlotaGO", int((current["Plataforma"] == "FLOTAGO").sum())],
        ["Conductores", current["Conductor"].nunique()], ["Equipos", current["Tracto"].nunique()], ["Fatiga", fatigue_total],
        ["Cumplimiento fatiga", f"{fatigue_pct:.1f}%"], ["Conductores reincidentes", recurrence["three_plus"]],
    ]
    elements += [_style_table(Table(kpi, colWidths=[5.0 * cm, 3.3 * cm]), 6.5), Spacer(1, 8), Image(str(p_evol), width=18 * cm, height=7.2 * cm), PageBreak()]
    elements += [Paragraph("TENDENCIAS", styles["TitleCustom"]), Paragraph("2. Evolución por tipo de alerta", styles["H2Custom"]), Image(str(p_types), width=18 * cm, height=8.4 * cm), PageBreak()]

    rows = [["Conductor", "Alertas", "Días", "Fatiga", "Celular", "Sem. ant.", "Tend."]]
    for _, r in drivers.iterrows():
        rows.append([Paragraph(truncate(r["Conductor"], 38), styles["SmallCustom"]), int(r["Alertas"]), int(r["Dias"]), int(r["Fatiga"]), int(r["Celular"]), int(r["Semana anterior"]), r["Tendencia"]])
    elements += [
        Paragraph("CONDUCTORES Y REINCIDENCIA", styles["TitleCustom"]),
        Paragraph("3. Conductores con mayor cantidad de alertas", styles["H2Custom"]),
        Image(str(p_drivers), width=18 * cm, height=7.5 * cm), Spacer(1, 8),
        _style_table(Table(rows, colWidths=[6.1 * cm, 1.5 * cm, 1.2 * cm, 1.3 * cm, 1.3 * cm, 1.6 * cm, 1.1 * cm]), 5.8),
        PageBreak(),
    ]

    fatigue_rows = [
        ["Indicador", "Resultado"], ["Fatiga", fatigue_total], ["Cumplen 15 minutos", fatigue_yes],
        ["No cumplen", fatigue_no], ["Pendientes", fatigue_pending], ["Cumplimiento", f"{fatigue_pct:.1f}%"],
    ]
    rec_rows = [
        ["Reincidencia", "Conductores"], ["1 alerta", recurrence["one"]], ["2 alertas", recurrence["two"]],
        ["3 o más alertas", recurrence["three_plus"]], ["% reincidencia", f"{recurrence['rate']:.1f}%"],
    ]
    elements += [
        Paragraph("FATIGA Y CONCLUSIONES", styles["TitleCustom"]),
        Paragraph("4. Cumplimiento y reincidencia", styles["H2Custom"]),
        Table([[_style_table(Table(fatigue_rows, colWidths=[5.5 * cm, 3.2 * cm]), 6.4), _style_table(Table(rec_rows, colWidths=[5.2 * cm, 3.2 * cm]), 6.4)]], colWidths=[9.1 * cm, 8.6 * cm], style=[("VALIGN", (0, 0), (-1, -1), "TOP")]),
        Spacer(1, 10), Paragraph(f"<b>Conclusión preventiva:</b> El transportista registra {total} alertas y una variación de {variation} respecto de la semana anterior. Se recomienda concentrar la gestión en los conductores reincidentes y en las alertas que presentan tendencia al alza.", styles["BodyCustom"]),
    ]
    doc.build(elements, onFirstPage=_footer, onLaterPages=_footer)
    return pdf


def generate_reports(data: pd.DataFrame, config: ReportConfig) -> tuple[bytes | None, bytes | None, dict]:
    data = filter_real_transportistas(data)
    current = _week_slice(data, config.week_start, config.week_end)
    if current.empty:
        raise ValueError("No existen alertas dentro del período seleccionado.")

    with tempfile.TemporaryDirectory(prefix="copec_reports_") as tmp:
        out = Path(tmp)
        global_pdf_bytes = None
        if config.include_global:
            global_path = generate_global_report(data, config, out)
            global_pdf_bytes = global_path.read_bytes()

        report_paths: list[Path] = []
        if config.include_transportistas:
            transportistas = sorted(current["Transportista"].dropna().unique())
            for i, transportista in enumerate(transportistas, start=1):
                report_paths.append(generate_transportista_report(data, transportista, config, out, i))
        else:
            transportistas = []

        zip_bytes = None
        if config.include_transportistas or config.include_global:
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                if config.include_global:
                    zf.writestr(Path(global_path).name, global_pdf_bytes)
                for path in report_paths:
                    zf.write(path, arcname=path.name)
            zip_bytes = buffer.getvalue()

    summary = {
        "total_alertas": len(current),
        "transportistas": len(transportistas),
        "conductores": current["Conductor"].nunique(),
        "equipos": current["Tracto"].nunique(),
        "fatiga": int((current["Tipo"] == "Fatiga").sum()),
    }
    return global_pdf_bytes, zip_bytes, summary
