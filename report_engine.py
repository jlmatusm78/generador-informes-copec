from __future__ import annotations

import io
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
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ALERT_WEIGHTS = {
    "Fatiga": 5, "Uso celular": 4, "Sin cinturón": 3, "Tapado cámara": 3,
    "Conductor fumando": 2, "Sin conductor": 2, "Cámara desalineada": 1, "Bostezo": 1,
}
GENERIC_TRANSPORTISTAS = ["COPEC", "NO ES COPEC", "PLANTA", "DESCONOCIDO", "OWL", "GPS", "PRUEBA", "SIN TRANSPORTISTA"]


@dataclass
class ReportConfig:
    period_start: date
    period_end: date
    report_mode: str = "Semanal"
    comparison_periods: int = 4
    include_global: bool = True
    include_transportistas: bool = True
    top_drivers: int = 12

    @property
    def week_start(self):
        return self.period_start

    @property
    def week_end(self):
        return self.period_end


def norm_text(value: object) -> str:
    text = str(value).upper().strip()
    text = "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", text)


def safe_filename(value: object, max_len: int = 90) -> str:
    text = "".join(c for c in unicodedata.normalize("NFD", str(value)) if unicodedata.category(c) != "Mn")
    text = re.sub(r"[^\w\s\-.&]", "", text)
    return (re.sub(r"\s+", " ", text).strip()[:max_len] or "SIN_NOMBRE")


def truncate(value: object, length: int) -> str:
    text = str(value)
    return text if len(text) <= length else text[: length - 1] + "…"


def wrap_label(value: object, width: int = 20) -> str:
    return "\n".join(textwrap.wrap(str(value), width=width))


def normalize_alert(value: object) -> str:
    s = str(value).lower()
    if any(x in s for x in ["cansancio", "fatiga", "somnolencia"]): return "Fatiga"
    if any(x in s for x in ["celular", "telefono", "teléfono"]): return "Uso celular"
    if "cintur" in s: return "Sin cinturón"
    if "desalinead" in s: return "Cámara desalineada"
    if "tapad" in s: return "Tapado cámara"
    if "bostezo" in s: return "Bostezo"
    if "fumando" in s: return "Conductor fumando"
    if "sin conductor" in s: return "Sin conductor"
    return str(value).strip()[:55] or "Sin clasificar"


def load_data(uploaded_file) -> pd.DataFrame:
    raw = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else Path(uploaded_file).read_bytes()
    excel = pd.ExcelFile(io.BytesIO(raw))
    frames = []
    for sheet in excel.sheet_names:
        name = sheet.strip().upper()
        if name in {"GUARDIAN", "FLOTAGO"}:
            df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet)
            df["Plataforma"] = "FLOTAGO" if "FLOTA" in name else "GUARDIAN"
            frames.append(df)
    if not frames:
        raise ValueError("El archivo debe contener hojas llamadas GUARDIAN y/o FLOTAGO.")
    data = pd.concat(frames, ignore_index=True)
    required = ["ID", "Fecha", "Transportista", "Conductor", "Tracto", "Incidente", "Plataforma", "Conductor se detine mínimo 15 minutos"]
    for col in required:
        if col not in data.columns: data[col] = np.nan
    data["Fecha"] = pd.to_datetime(data["Fecha"], errors="coerce")
    data = data.dropna(subset=["Fecha"]).copy()
    data["Transportista"] = data["Transportista"].fillna("SIN TRANSPORTISTA").astype(str).str.strip()
    data["Conductor"] = data["Conductor"].fillna("SIN CONDUCTOR").astype(str).str.strip()
    data["Tracto"] = data["Tracto"].fillna("S/I").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    data["Incidente"] = data["Incidente"].fillna("Sin clasificar").astype(str).str.strip()
    data["Plataforma"] = data["Plataforma"].fillna("SIN PLATAFORMA").astype(str).str.upper().str.strip()
    data["Tipo"] = data["Incidente"].apply(normalize_alert)
    data["FechaDia"] = data["Fecha"].dt.date
    data["Mes"] = data["Fecha"].dt.to_period("M")
    data["Puntaje"] = data["Tipo"].map(ALERT_WEIGHTS).fillna(1).astype(int)
    return data.sort_values("Fecha").reset_index(drop=True)


def filter_real_transportistas(data: pd.DataFrame) -> pd.DataFrame:
    return data[data["Transportista"].map(lambda x: not any(g in norm_text(x) for g in GENERIC_TRANSPORTISTAS))].copy()


def detect_default_week(data: pd.DataFrame) -> tuple[date, date]:
    max_date = data["Fecha"].max().date()
    end = max_date - timedelta(days=(max_date.weekday() + 1) % 7)
    return end - timedelta(days=6), end


def detect_default_month(data: pd.DataFrame) -> tuple[date, date]:
    max_date = data["Fecha"].max()
    p = max_date.to_period("M")
    return p.start_time.date(), p.end_time.date()


def _period_slice(data: pd.DataFrame, start: date, end: date) -> pd.DataFrame:
    return data[(data["Fecha"] >= pd.Timestamp(start)) & (data["Fecha"] <= pd.Timestamp(datetime.combine(end, datetime.max.time())))].copy()


def comparison_ranges(config: ReportConfig):
    ranges = []
    if config.report_mode.lower().startswith("mens"):
        current = pd.Period(config.period_start, freq="M")
        for p in pd.period_range(current - (config.comparison_periods - 1), current, freq="M"):
            ranges.append((p.strftime("%b-%Y"), p.start_time, p.end_time))
    else:
        first = config.period_start - timedelta(days=7 * (config.comparison_periods - 1))
        for i in range(config.comparison_periods):
            s = first + timedelta(days=7 * i); e = s + timedelta(days=6)
            ranges.append((f"{s:%d-%m} / {e:%d-%m}", pd.Timestamp(s), pd.Timestamp(datetime.combine(e, datetime.max.time()))))
    return ranges


def pct_value(current: float, previous: float) -> float:
    if previous == 0: return 100.0 if current > 0 else 0.0
    return (current - previous) / previous * 100


def pct_change(current: float, previous: float) -> str:
    return f"{pct_value(current, previous):+.1f}%"


def trend_symbol(current: float, previous: float) -> str:
    return "↑" if current > previous else ("↓" if current < previous else "↔")


def risk_level(score: float, alerts: int, fatigue: int) -> str:
    if fatigue >= 3 or score >= 25 or alerts >= 12: return "CRÍTICO"
    if fatigue >= 1 or score >= 12 or alerts >= 6: return "ALTO"
    if score >= 6 or alerts >= 3: return "MEDIO"
    return "BAJO"


def compliance_summary(df: pd.DataFrame):
    fatigue = df[df["Tipo"].eq("Fatiga")]
    total = len(fatigue)
    if not total: return 0, 0, 0, 0, 0.0
    v = fatigue["Conductor se detine mínimo 15 minutos"].fillna("PENDIENTE").astype(str).str.upper().str.strip()
    yes, no = int((v == "SI").sum()), int((v == "NO").sum())
    return total, yes, no, total - yes - no, yes / total * 100


def recurrence_summary(df: pd.DataFrame):
    counts = df.groupby("Conductor").size() if not df.empty else pd.Series(dtype=int)
    total = len(counts); three = int((counts >= 3).sum())
    return {"one": int((counts == 1).sum()), "two": int((counts == 2).sum()), "three_plus": three, "total": total, "rate": three / total * 100 if total else 0.0}


def detailed_driver_ranking(current: pd.DataFrame, previous: pd.DataFrame, topn=12, include_transportista=False):
    group = ["Conductor", "Transportista"] if include_transportista else ["Conductor"]
    cols = group + ["Alertas", "Puntaje", "Fatiga", "Celular", "Cámaras", "Anterior", "Variación", "Tendencia", "Nivel"]
    if current.empty: return pd.DataFrame(columns=cols)
    rank = current.groupby(group).agg(
        Alertas=("ID", "count"), Puntaje=("Puntaje", "sum"),
        Fatiga=("Tipo", lambda s: int((s == "Fatiga").sum())),
        Celular=("Tipo", lambda s: int((s == "Uso celular").sum())),
        Cámaras=("Tipo", lambda s: int(s.isin(["Tapado cámara", "Cámara desalineada"]).sum())),
        Dias=("FechaDia", "nunique"),
    ).reset_index()
    prev = previous.groupby("Conductor").size()
    rank["Anterior"] = rank["Conductor"].map(prev).fillna(0).astype(int)
    rank["Variación"] = [pct_change(a, b) for a, b in zip(rank["Alertas"], rank["Anterior"])]
    rank["Tendencia"] = [trend_symbol(a, b) for a, b in zip(rank["Alertas"], rank["Anterior"])]
    rank["Nivel"] = [risk_level(s, a, f) for s, a, f in zip(rank["Puntaje"], rank["Alertas"], rank["Fatiga"])]
    return rank.sort_values(["Puntaje", "Fatiga", "Alertas"], ascending=False).head(topn)


def _styles():
    s = getSampleStyleSheet()
    s.add(ParagraphStyle(name="BodyCustom", parent=s["BodyText"], fontSize=8.1, leading=10.4))
    s.add(ParagraphStyle(name="SmallCustom", parent=s["BodyText"], fontSize=6.0, leading=7.3))
    s.add(ParagraphStyle(name="TitleCustom", parent=s["Title"], fontSize=16, leading=19, textColor=colors.HexColor("#1F4E79")))
    s.add(ParagraphStyle(name="H2Custom", parent=s["Heading2"], fontSize=10.5, leading=13, textColor=colors.HexColor("#1F4E79")))
    return s


def _style_table(table: Table, font_size=6.1):
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E79")), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,-1), font_size),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#C9D6E2")), ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")]),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("ALIGN", (1,1), (-1,-1), "CENTER"),
        ("LEFTPADDING", (0,0), (-1,-1), 2.3), ("RIGHTPADDING", (0,0), (-1,-1), 2.3), ("TOPPADDING", (0,0), (-1,-1), 2), ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ])); return table


def _footer(canvas, doc):
    canvas.saveState(); canvas.setFont("Helvetica", 7); canvas.setFillColor(colors.grey)
    canvas.drawString(1.4*cm, .8*cm, "Torre de Control COPEC - Informe automático")
    canvas.drawRightString(A4[0]-1.4*cm, .8*cm, f"Página {doc.page}"); canvas.restoreState()


def _plot_line(labels, series, title, ylabel, path):
    plt.figure(figsize=(8.6, 3.8))
    for label, values in series.items(): plt.plot(labels, values, marker="o", linewidth=2, label=label)
    plt.title(title); plt.ylabel(ylabel); plt.grid(True, alpha=.25); plt.xticks(rotation=15)
    if len(series) > 1: plt.legend(fontsize=7, ncol=3, loc="upper center", bbox_to_anchor=(.5,-.2))
    plt.tight_layout(); plt.savefig(path, dpi=140, bbox_inches="tight"); plt.close()


def _plot_bar(labels, values: Iterable[int], title, xlabel, path):
    plt.figure(figsize=(8.4, 4.2)); plt.barh(labels[::-1], list(values)[::-1]); plt.title(title); plt.xlabel(xlabel)
    plt.tight_layout(); plt.savefig(path, dpi=140); plt.close()


def _metrics(data, ranges):
    top_types = data["Tipo"].value_counts().head(4).index.tolist()
    totals, by_type = [], {t: [] for t in top_types}
    for _, s, e in ranges:
        x = data[(data["Fecha"] >= s) & (data["Fecha"] <= e)]
        totals.append(len(x))
        for t in top_types: by_type[t].append(int((x["Tipo"] == t).sum()))
    return totals, by_type


def _insights(current, previous, rank_trans=None, drivers=None):
    total, prev = len(current), len(previous)
    insights = [f"Las alertas presentan una variación de {pct_change(total, prev)} respecto del período anterior."]
    types = current["Tipo"].value_counts()
    if len(types): insights.append(f"El tipo más frecuente es {types.index[0]} con {int(types.iloc[0])} eventos ({types.iloc[0]/max(total,1)*100:.1f}% del total).")
    if drivers is not None and len(drivers):
        concentration = drivers.head(10)["Alertas"].sum() / max(total,1) * 100
        insights.append(f"Los 10 conductores más críticos concentran {concentration:.1f}% de las alertas del período.")
    if rank_trans is not None and len(rank_trans): insights.append(f"{rank_trans.iloc[0]['Transportista']} lidera el ranking con {int(rank_trans.iloc[0]['Alertas'])} alertas.")
    ft, fy, fn, fp, pct = compliance_summary(current)
    if ft: insights.append(f"El cumplimiento registrado del protocolo de fatiga alcanza {pct:.1f}% ({fn} no cumplimientos y {fp} pendientes).")
    return insights[:5]


def _rank_transportistas(current, previous):
    r = current.groupby("Transportista").agg(Alertas=("ID","count"), Conductores=("Conductor","nunique"), Equipos=("Tracto","nunique"), Fatiga=("Tipo",lambda s:int((s=="Fatiga").sum())), Puntaje=("Puntaje","sum")).reset_index()
    p = previous.groupby("Transportista").size()
    r["Anterior"] = r["Transportista"].map(p).fillna(0).astype(int)
    r["Variación"] = [pct_change(a,b) for a,b in zip(r.Alertas,r.Anterior)]
    r["Tendencia"] = [trend_symbol(a,b) for a,b in zip(r.Alertas,r.Anterior)]
    r["Estado"] = [risk_level(s,a,f) for s,a,f in zip(r.Puntaje,r.Alertas,r.Fatiga)]
    return r.sort_values(["Puntaje","Fatiga","Alertas"], ascending=False)


def generate_global_report(data, config, output_dir):
    styles = _styles(); current = _period_slice(data, config.period_start, config.period_end)
    ranges = comparison_ranges(config); labels=[x[0] for x in ranges]; totals, by_type = _metrics(data, ranges)
    previous = data[(data["Fecha"]>=ranges[-2][1]) & (data["Fecha"]<=ranges[-2][2])] if len(ranges)>1 else current.iloc[0:0]
    rank_t = _rank_transportistas(current, previous); drivers = detailed_driver_ranking(current, previous, config.top_drivers, True)
    chart_dir=output_dir/"charts_global"; chart_dir.mkdir(exist_ok=True)
    p_total=chart_dir/"total.png"; p_type=chart_dir/"tipos.png"; p_trans=chart_dir/"transportistas.png"; p_drv=chart_dir/"conductores.png"
    unit="mensual" if config.report_mode.lower().startswith("mens") else "semanal"
    _plot_line(labels,{"Alertas":totals},f"Evolución {unit} de alertas","Alertas",p_total)
    _plot_line(labels,by_type or {"Sin datos":[0]*len(labels)},f"Tendencia {unit} por tipo de alerta","Alertas",p_type)
    _plot_bar([wrap_label(x,24) for x in rank_t.head(10).Transportista],rank_t.head(10).Alertas,"Transportistas prioritarios","Alertas",p_trans)
    _plot_bar([wrap_label(x,27) for x in drivers.head(10).Conductor],drivers.head(10).Puntaje,"Conductores por puntaje de criticidad","Puntaje",p_drv)
    total=len(current); fatigue=compliance_summary(current); rec=recurrence_summary(current)
    pdf=output_dir/f"00_Informe_Global_COPEC_{config.report_mode}_{config.period_start:%d%m}_{config.period_end:%d%m%Y}.pdf"
    doc=SimpleDocTemplate(str(pdf),pagesize=A4,rightMargin=1.05*cm,leftMargin=1.05*cm,topMargin=.9*cm,bottomMargin=1.2*cm)
    title=f"INFORME {config.report_mode.upper()} GLOBAL COPEC"
    kpi=[["Indicador","Resultado"],["Total alertas",total],["Variación",pct_change(total,len(previous))],["Promedio diario",f"{total/max((config.period_end-config.period_start).days+1,1):.1f}"],["Transportistas",current.Transportista.nunique()],["Conductores",current.Conductor.nunique()],["Equipos",current.Tracto.nunique()],["Fatiga",fatigue[0]],["Cumplimiento fatiga",f"{fatigue[4]:.1f}%"],["Reincidentes 3+",rec['three_plus']]]
    sem=[["Prioridad","Resultado"],["Transportistas críticos",int((rank_t.Estado=="CRÍTICO").sum())],["Conductores críticos",int((drivers.Nivel=="CRÍTICO").sum())],["No cumplimiento fatiga",fatigue[2]],["Pendientes fatiga",fatigue[3]],["Concentración Top 10",f"{drivers.head(10).Alertas.sum()/max(total,1)*100:.1f}%"]]
    els=[Paragraph(title,styles["TitleCustom"]),Paragraph(f"Período: {config.period_start:%d-%m-%Y} al {config.period_end:%d-%m-%Y} · Guardian y FlotaGO",styles["BodyCustom"]),Spacer(1,6),Paragraph("1. Panel ejecutivo",styles["H2Custom"]),Table([[_style_table(Table(kpi,colWidths=[4.3*cm,2.7*cm]),6.3),_style_table(Table(sem,colWidths=[4.5*cm,2.7*cm]),6.3)]],colWidths=[8.5*cm,8.5*cm],style=[("VALIGN",(0,0),(-1,-1),"TOP")]),Spacer(1,5),Image(str(p_total),width=17.8*cm,height=7.3*cm)]
    insights=_insights(current,previous,rank_t,drivers)
    els += [Spacer(1,4),Paragraph("Decisiones rápidas",styles["H2Custom"]),_style_table(Table([[str(i+1),Paragraph(x,styles["SmallCustom"])] for i,x in enumerate(insights)],colWidths=[.7*cm,16.3*cm]),6.0),PageBreak()]
    rows=[["Transportista","Alertas","Ant.","Var.","Fatiga","Reinc.","Estado"]]
    for _,r in rank_t.head(12).iterrows():
        rr=recurrence_summary(current[current.Transportista==r.Transportista])
        rows.append([Paragraph(truncate(r.Transportista,31),styles["SmallCustom"]),int(r.Alertas),int(r.Anterior),r.Variación,int(r.Fatiga),rr['three_plus'],r.Estado])
    els += [Paragraph("TRANSPORTISTAS Y TENDENCIAS",styles["TitleCustom"]),Paragraph("2. Dónde concentrar la gestión",styles["H2Custom"]),Image(str(p_trans),width=17.8*cm,height=7.0*cm),Spacer(1,5),_style_table(Table(rows,colWidths=[5.0*cm,1.3*cm,1.2*cm,1.4*cm,1.2*cm,1.3*cm,1.7*cm]),5.6),Spacer(1,5),Image(str(p_type),width=17.8*cm,height=6.3*cm),PageBreak()]
    drows=[["Conductor","Empresa","Alertas","Puntos","Fatiga","Cel.","Cám.","Var.","Nivel"]]
    for _,r in drivers.iterrows(): drows.append([Paragraph(truncate(r.Conductor,27),styles["SmallCustom"]),Paragraph(truncate(r.Transportista,23),styles["SmallCustom"]),int(r.Alertas),int(r.Puntaje),int(r.Fatiga),int(r.Celular),int(r["Cámaras"]),r.Variación,r.Nivel])
    els += [Paragraph("CONDUCTORES CRÍTICOS",styles["TitleCustom"]),Paragraph("3. Ranking ponderado por severidad",styles["H2Custom"]),Image(str(p_drv),width=17.8*cm,height=7.0*cm),Spacer(1,5),_style_table(Table(drows,colWidths=[4.0*cm,3.3*cm,1.1*cm,1.1*cm,1.0*cm,1.0*cm,1.0*cm,1.3*cm,1.5*cm]),5.2),Spacer(1,5),Paragraph("El puntaje asigna mayor peso a fatiga, uso de celular, falta de cinturón y tapado de cámara. Esto permite priorizar riesgo y no solo volumen.",styles["BodyCustom"]),PageBreak()]
    actions=[["Prioridad","Hallazgo","Acción sugerida"]]
    for _,r in rank_t.head(5).iterrows(): actions.append([r.Estado,Paragraph(truncate(f"{r.Transportista}: {r.Alertas} alertas, variación {r.Variación}",60),styles["SmallCustom"]),Paragraph("Solicitar análisis causal, responsables y seguimiento del próximo período.",styles["SmallCustom"])])
    actions += [["ALTA","Fatiga sin cumplimiento o pendiente","Revisar gestión, descansos y trazabilidad de llamados."],["ALTA","Conductores críticos recurrentes","Aplicar retroalimentación individual y verificar reincidencia."],["MEDIA","Alertas de cámara","Separar revisión técnica de posible manipulación."]]
    els += [Paragraph("CONCLUSIONES Y PLAN DE ACCIÓN",styles["TitleCustom"]),Paragraph("4. Priorización ejecutiva",styles["H2Custom"]),_style_table(Table(actions,colWidths=[1.7*cm,7.2*cm,8.1*cm]),5.8),Spacer(1,8),Paragraph("Criterio recomendado: intervenir primero las combinaciones de alta severidad, tendencia creciente y reincidencia. Mantener seguimiento diferenciado para alertas técnicas de cámara.",styles["BodyCustom"])]
    doc.build(els,onFirstPage=_footer,onLaterPages=_footer); return pdf


def generate_transportista_report(data, transportista, config, output_dir, index):
    styles=_styles(); all_t=data[data.Transportista==transportista].copy(); current=_period_slice(all_t,config.period_start,config.period_end)
    ranges=comparison_ranges(config); labels=[x[0] for x in ranges]; totals,by_type=_metrics(all_t,ranges)
    previous=all_t[(all_t.Fecha>=ranges[-2][1])&(all_t.Fecha<=ranges[-2][2])] if len(ranges)>1 else current.iloc[0:0]
    drivers=detailed_driver_ranking(current,previous,config.top_drivers,False); fatigue=compliance_summary(current); rec=recurrence_summary(current)
    chart=output_dir/f"charts_{index:02d}"; chart.mkdir(exist_ok=True); p1=chart/"evol.png"; p2=chart/"tipos.png"; p3=chart/"drivers.png"
    unit="mensual" if config.report_mode.lower().startswith("mens") else "semanal"
    _plot_line(labels,{"Alertas":totals},f"Evolución {unit}","Alertas",p1); _plot_line(labels,by_type or {"Sin datos":[0]*len(labels)},"Tendencia por tipo","Alertas",p2); _plot_bar([wrap_label(x,26) for x in drivers.head(10).Conductor],drivers.head(10).Puntaje,"Ranking de criticidad","Puntaje",p3)
    pdf=output_dir/f"{index:02d}_Informe_{config.report_mode}_{safe_filename(transportista)}_{config.period_start:%d%m}_{config.period_end:%d%m%Y}.pdf"
    doc=SimpleDocTemplate(str(pdf),pagesize=A4,rightMargin=1.05*cm,leftMargin=1.05*cm,topMargin=.9*cm,bottomMargin=1.2*cm)
    total=len(current); kpi=[["Indicador","Resultado"],["Alertas",total],["Período anterior",len(previous)],["Variación",pct_change(total,len(previous))],["Promedio diario",f"{total/max((config.period_end-config.period_start).days+1,1):.1f}"],["Conductores",current.Conductor.nunique()],["Equipos",current.Tracto.nunique()],["Fatiga",fatigue[0]],["Cumplimiento",f"{fatigue[4]:.1f}%"],["Reincidentes 3+",rec['three_plus']]]
    insights=_insights(current,previous,None,drivers)
    els=[Paragraph(f"INFORME {config.report_mode.upper()} DE ALERTAS",styles["TitleCustom"]),Paragraph(f"Transportista: {transportista}<br/>Período: {config.period_start:%d-%m-%Y} al {config.period_end:%d-%m-%Y}",styles["BodyCustom"]),Spacer(1,6),Paragraph("1. Resumen ejecutivo",styles["H2Custom"]),Table([[_style_table(Table(kpi,colWidths=[4.4*cm,2.8*cm]),6.3),Image(str(p1),width=9.4*cm,height=6.1*cm)]],colWidths=[7.6*cm,9.6*cm],style=[("VALIGN",(0,0),(-1,-1),"TOP")]),Spacer(1,5),_style_table(Table([[str(i+1),Paragraph(x,styles["SmallCustom"])] for i,x in enumerate(insights)],colWidths=[.7*cm,16.3*cm]),6.0),Spacer(1,5),Image(str(p2),width=17.8*cm,height=6.6*cm),PageBreak()]
    rows=[["Conductor","Alertas","Puntos","Fatiga","Celular","Cámaras","Ant.","Var.","Nivel"]]
    for _,r in drivers.iterrows(): rows.append([Paragraph(truncate(r.Conductor,34),styles["SmallCustom"]),int(r.Alertas),int(r.Puntaje),int(r.Fatiga),int(r.Celular),int(r["Cámaras"]),int(r.Anterior),r.Variación,r.Nivel])
    els += [Paragraph("CONDUCTORES Y TENDENCIAS",styles["TitleCustom"]),Paragraph("2. Ranking detallado por riesgo",styles["H2Custom"]),Image(str(p3),width=17.8*cm,height=7.2*cm),Spacer(1,5),_style_table(Table(rows,colWidths=[5.2*cm,1.0*cm,1.0*cm,1.0*cm,1.0*cm,1.1*cm,1.0*cm,1.3*cm,1.4*cm]),5.2),Spacer(1,5),Paragraph("La tendencia compara cada conductor con el período inmediatamente anterior. El puntaje pondera la severidad de cada tipo de alerta.",styles["BodyCustom"]),PageBreak()]
    types=current.Tipo.value_counts().head(8); type_rows=[["Tipo de alerta","Cantidad","% total"]]+[[t,int(v),f"{v/max(total,1)*100:.1f}%"] for t,v in types.items()]
    actions=[["Prioridad","Acción recomendada"],["Alta","Gestionar conductores críticos y reincidentes con seguimiento individual."],["Alta","Revisar todos los no cumplimientos y pendientes de fatiga."],["Media","Separar fallas técnicas de cámara de posibles manipulaciones."],["Media","Comparar el resultado en el siguiente período y verificar reducción."]]
    els += [Paragraph("COMPOSICIÓN Y ACCIONES",styles["TitleCustom"]),Paragraph("3. Tipos de alerta, fatiga y plan",styles["H2Custom"]),Table([[_style_table(Table(type_rows,colWidths=[4.8*cm,1.8*cm,1.8*cm]),6.0),_style_table(Table([["Fatiga","Resultado"],["Total",fatigue[0]],["Cumple",fatigue[1]],["No cumple",fatigue[2]],["Pendiente",fatigue[3]],["Cumplimiento",f"{fatigue[4]:.1f}%"]],colWidths=[4.5*cm,2.3*cm]),6.0)]],colWidths=[9.0*cm,8.0*cm],style=[("VALIGN",(0,0),(-1,-1),"TOP")]),Spacer(1,8),_style_table(Table(actions,colWidths=[2.0*cm,15.0*cm]),6.0)]
    doc.build(els,onFirstPage=_footer,onLaterPages=_footer); return pdf


def generate_reports(data, config):
    data=filter_real_transportistas(data); current=_period_slice(data,config.period_start,config.period_end)
    if current.empty: raise ValueError("No existen alertas dentro del período seleccionado.")
    with tempfile.TemporaryDirectory(prefix="copec_reports_") as tmp:
        out=Path(tmp); global_bytes=None; report_paths=[]; global_path=None
        if config.include_global:
            global_path=generate_global_report(data,config,out); global_bytes=global_path.read_bytes()
        transportistas=sorted(current.Transportista.dropna().unique()) if config.include_transportistas else []
        for i,t in enumerate(transportistas,1): report_paths.append(generate_transportista_report(data,t,config,out,i))
        buffer=io.BytesIO()
        with zipfile.ZipFile(buffer,"w",zipfile.ZIP_DEFLATED) as zf:
            if global_path: zf.write(global_path,arcname=global_path.name)
            for p in report_paths: zf.write(p,arcname=p.name)
        zip_bytes=buffer.getvalue()
    return global_bytes,zip_bytes,{"total_alertas":len(current),"transportistas":len(transportistas),"conductores":current.Conductor.nunique(),"equipos":current.Tracto.nunique(),"fatiga":int((current.Tipo=="Fatiga").sum()),"modo":config.report_mode}
