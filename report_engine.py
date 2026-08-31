from __future__ import annotations

import io, re, tempfile, textwrap, unicodedata, zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, LongTable, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from alert_aliases import SENSOR_BLOCKED, SENSOR_MISALIGNED, normalize_alert_name

RISK_WEIGHTS = {"Fatiga sin cumplimiento": 10, "Fatiga pendiente": 8, "Sin cinturón": 6, "Conductor fumando": 6, "Fatiga con cumplimiento": 3}
GENERIC_TRANSPORTISTAS = ["COPEC", "NO ES COPEC", "PLANTA", "DESCONOCIDO", "OWL", "GPS", "PRUEBA", "SIN TRANSPORTISTA"]
REPORT_ENGINE_VERSION = "2026.08.31.2"
REPORT_LOGO = Path(__file__).resolve().parent / "assets" / "logo_copec_90.png"


@dataclass
class ReportConfig:
    period_start: date
    period_end: date
    report_mode: str = "Semanal"
    comparison_periods: int = 4
    include_global: bool = True
    include_transportistas: bool = True
    top_drivers: int = 10

    @property
    def week_start(self): return self.period_start

    @property
    def week_end(self): return self.period_end


def norm_text(value):
    text = str(value).upper().strip()
    text = "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", text)


def safe_filename(value, max_len=90):
    text = "".join(c for c in unicodedata.normalize("NFD", str(value)) if unicodedata.category(c) != "Mn")
    text = re.sub(r"[^\w\s\-.&]", "", text)
    return (re.sub(r"\s+", " ", text).strip()[:max_len] or "SIN_NOMBRE")


def truncate(value, length):
    text = str(value)
    return text if len(text) <= length else text[:length-1] + "…"


def wrap_label(value, width=20): return "\n".join(textwrap.wrap(str(value), width=width))


def normalize_alert(value):
    canonical = normalize_alert_name(str(value).strip()); text = canonical.lower()
    if any(x in text for x in ["cansancio", "fatiga", "somnolencia"]): return "Fatiga"
    if any(x in text for x in ["celular", "telefono", "teléfono"]): return "Uso celular"
    if "cintur" in text: return "Sin cinturón"
    if "desalinead" in text: return SENSOR_MISALIGNED
    if "tapad" in text: return SENSOR_BLOCKED
    if "bostezo" in text: return "Bostezo"
    if "fumando" in text: return "Conductor fumando"
    if "sin conductor" in text: return "Sin conductor"
    return canonical[:55] or "Sin clasificar"


def normalize_compliance(value):
    text = norm_text(value)
    if text in {"SI", "YES", "CUMPLE"}: return "SI"
    if text in {"NO", "NO CUMPLE"}: return "NO"
    return "PENDIENTE"


def load_data(uploaded_file):
    raw = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else Path(uploaded_file).read_bytes()
    excel = pd.ExcelFile(io.BytesIO(raw)); frames = []
    for sheet in excel.sheet_names:
        name = sheet.strip().upper()
        if name in {"GUARDIAN", "FLOTAGO"}:
            frame = pd.read_excel(io.BytesIO(raw), sheet_name=sheet)
            frame["Plataforma"] = "FLOTAGO" if "FLOTA" in name else "GUARDIAN"; frames.append(frame)
    if not frames: raise ValueError("El archivo debe contener hojas llamadas GUARDIAN y/o FLOTAGO.")
    data = pd.concat(frames, ignore_index=True)
    required = ["ID","Fecha","Transportista","Conductor","Tracto","Patente","Incidente","Plataforma","Fecha/Hora de evento","Fecha/Hora de gestion","Tiempo respuesta evento. (No ingresar datos)","Velocidad (sólo número)","Monitor","Observaciones","Estado","Conductor se detine mínimo 15 minutos"]
    for col in required:
        if col not in data.columns: data[col] = np.nan
    data["Fecha"] = pd.to_datetime(data["Fecha"], errors="coerce", dayfirst=True); data = data.dropna(subset=["Fecha"]).copy()
    data["Transportista"] = data["Transportista"].fillna("SIN TRANSPORTISTA").astype(str).str.strip()
    data["Conductor"] = data["Conductor"].fillna("SIN CONDUCTOR").astype(str).str.strip()
    data["Tracto"] = data["Tracto"].fillna("S/I").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    data["Patente"] = data["Patente"].fillna("S/I").astype(str).str.strip()
    data["Incidente"] = data["Incidente"].fillna("Sin clasificar").astype(str).str.strip()
    data["Plataforma"] = data["Plataforma"].fillna("SIN PLATAFORMA").astype(str).str.upper().str.strip()
    data["Tipo"] = data["Incidente"].apply(normalize_alert)
    sensor_rows = data["Tipo"].isin({SENSOR_MISALIGNED, SENSOR_BLOCKED})
    data.loc[sensor_rows, "Incidente"] = data.loc[sensor_rows, "Tipo"]
    data["CumplimientoFatiga"] = data["Conductor se detine mínimo 15 minutos"].apply(normalize_compliance)
    data["FechaDia"] = data["Fecha"].dt.date; data["Mes"] = data["Fecha"].dt.to_period("M")
    return data.sort_values("Fecha").reset_index(drop=True)


def filter_real_transportistas(data):
    return data[data["Transportista"].map(lambda x: not any(g in norm_text(x) for g in GENERIC_TRANSPORTISTAS))].copy()


def detect_default_week(data):
    max_date = data["Fecha"].max().date(); end = max_date - timedelta(days=(max_date.weekday()+1)%7)
    return end-timedelta(days=6), end


def detect_default_month(data):
    period = data["Fecha"].max().to_period("M"); return period.start_time.date(), period.end_time.date()


def _period_slice(data, start, end):
    return data[(data["Fecha"] >= pd.Timestamp(start)) & (data["Fecha"] <= pd.Timestamp(datetime.combine(end, datetime.max.time())))].copy()


def comparison_ranges(config):
    ranges = []
    if config.report_mode.lower().startswith("mens"):
        current = pd.Period(config.period_start, freq="M")
        for period in pd.period_range(current-(config.comparison_periods-1), current, freq="M"):
            ranges.append((period.strftime("%b-%Y"), period.start_time, period.end_time))
    else:
        first = config.period_start-timedelta(days=7*(config.comparison_periods-1))
        for i in range(config.comparison_periods):
            start = first+timedelta(days=7*i); end = start+timedelta(days=6)
            ranges.append((f"{start:%d-%m} / {end:%d-%m}", pd.Timestamp(start), pd.Timestamp(datetime.combine(end, datetime.max.time()))))
    return ranges


def pct_value(current, previous):
    if previous == 0: return 100.0 if current > 0 else 0.0
    return (current-previous)/previous*100


def pct_change(current, previous): return f"{pct_value(current, previous):+.1f}%"


def compliance_summary(df):
    fatigue = df[df["Tipo"].eq("Fatiga")]; total = len(fatigue)
    if not total: return 0,0,0,0,0.0
    yes = int((fatigue["CumplimientoFatiga"]=="SI").sum()); no = int((fatigue["CumplimientoFatiga"]=="NO").sum()); pending = total-yes-no
    return total,yes,no,pending,yes/total*100


def recurrence_summary(df):
    counts = df.groupby("Conductor").size() if not df.empty else pd.Series(dtype=int); total = len(counts); recurrent = int((counts>=2).sum())
    return {"one":int((counts==1).sum()),"two":int((counts==2).sum()),"three_plus":int((counts>=3).sum()),"recurrent":recurrent,"total":total,"rate":recurrent/total*100 if total else 0.0}


def operational_risk_summary(df):
    fatigue = df[df["Tipo"].eq("Fatiga")]
    no = int((fatigue["CumplimientoFatiga"]=="NO").sum()); pending = int((fatigue["CumplimientoFatiga"]=="PENDIENTE").sum()); yes = int((fatigue["CumplimientoFatiga"]=="SI").sum())
    belt = int((df["Tipo"]=="Sin cinturón").sum()); smoking = int((df["Tipo"]=="Conductor fumando").sum())
    points = no*10+pending*8+belt*6+smoking*6+yes*3
    repeated = int((fatigue[fatigue["CumplimientoFatiga"]=="NO"].groupby("Conductor").size()>=2).sum())
    if points >= 50 or repeated: level, rule = "CRÍTICO", ("50 puntos o más" if points>=50 else "reincidencia de fatiga sin cumplimiento")
    elif points >= 25: level, rule = "ALTO", "entre 25 y 49 puntos"
    elif points >= 10: level, rule = "MEDIO", "entre 10 y 24 puntos"
    else: level, rule = "BAJO", "menos de 10 puntos"
    equation = f"{no}×10 (fatiga sin cumplimiento) + {pending}×8 (fatiga pendiente) + {belt}×6 (sin cinturón) + {smoking}×6 (conductor fumando) + {yes}×3 (fatiga con cumplimiento) = {points} puntos."
    return {"fatigue_no":no,"fatigue_pending":pending,"fatigue_yes":yes,"seatbelt":belt,"smoking":smoking,"points":points,"level":level,"rule":rule,"repeated_fatigue_no_drivers":repeated,"equation":equation}


def detailed_driver_ranking(current, previous, topn=10, include_transportista=False):
    group = ["Conductor","Transportista"] if include_transportista else ["Conductor"]; rows=[]
    if current.empty: return pd.DataFrame()
    previous_counts = previous.groupby("Conductor").size()
    for keys, frame in current.groupby(group, dropna=False):
        keys = keys if isinstance(keys,tuple) else (keys,); risk=operational_risk_summary(frame); alerts=len(frame)
        row={group[i]:keys[i] for i in range(len(group))}; row.update({"Alertas":alerts,"Reincidencias":max(alerts-1,0),"Fatiga sin cumplir":risk["fatigue_no"],"Fatiga pendiente":risk["fatigue_pending"],"Sin cinturón":risk["seatbelt"],"Fumando":risk["smoking"],"Riesgo":risk["points"],"Nivel":risk["level"],"Anterior":int(previous_counts.get(keys[0],0))})
        row["Variación"] = pct_change(alerts,row["Anterior"]); rows.append(row)
    ranking=pd.DataFrame(rows); ranking["Prioritarios"]=ranking["Sin cinturón"]+ranking["Fumando"]
    return ranking.sort_values(["Fatiga sin cumplir","Reincidencias","Prioritarios","Alertas"],ascending=False).head(max(int(topn),1)).drop(columns=["Prioritarios"])


def _styles():
    styles=getSampleStyleSheet(); styles.add(ParagraphStyle(name="BodyCustom",parent=styles["BodyText"],fontSize=8.1,leading=10.4)); styles.add(ParagraphStyle(name="SmallCustom",parent=styles["BodyText"],fontSize=6.0,leading=7.3)); styles.add(ParagraphStyle(name="TinyCustom",parent=styles["BodyText"],fontSize=5.0,leading=6.0)); styles.add(ParagraphStyle(name="TitleCustom",parent=styles["Title"],fontSize=16,leading=19,textColor=colors.HexColor("#1F4E79"))); styles.add(ParagraphStyle(name="H2Custom",parent=styles["Heading2"],fontSize=10.5,leading=13,textColor=colors.HexColor("#1F4E79"))); return styles


def _style_table(table, font_size=6.1):
    table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E79")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),font_size),("GRID",(0,0),(-1,-1),.3,colors.HexColor("#C9D6E2")),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F7FAFC")]),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(1,1),(-1,-1),"CENTER"),("LEFTPADDING",(0,0),(-1,-1),2.3),("RIGHTPADDING",(0,0),(-1,-1),2.3),("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2)])); return table


def _page_branding(canvas, doc):
    canvas.saveState()
    page_width, page_height = doc.pagesize
    if REPORT_LOGO.exists():
        logo_width = 4.8 * cm
        logo_height = logo_width * 655 / 2286
        canvas.drawImage(
            str(REPORT_LOGO),
            1.05 * cm,
            page_height - 0.65 * cm - logo_height,
            width=logo_width,
            height=logo_height,
            preserveAspectRatio=True,
            mask="auto",
        )
    canvas.setStrokeColor(colors.HexColor("#E31B16"))
    canvas.setLineWidth(0.6)
    canvas.line(1.05 * cm, page_height - 2.05 * cm, page_width - 1.05 * cm, page_height - 2.05 * cm)
    canvas.setFont("Helvetica",7)
    canvas.setFillColor(colors.grey)
    canvas.drawString(1.4*cm,.8*cm,"Torre de Control COPEC - Informe automático")
    canvas.drawRightString(page_width-1.4*cm,.8*cm,f"Página {doc.page}")
    canvas.restoreState()


def _plot_line(labels,series,title,ylabel,path):
    plt.figure(figsize=(8.6,3.8))
    for label,values in series.items(): plt.plot(labels,values,marker="o",linewidth=2,label=label)
    plt.title(title); plt.ylabel(ylabel); plt.grid(True,alpha=.25); plt.xticks(rotation=15)
    if len(series)>1: plt.legend(fontsize=7,ncol=3,loc="upper center",bbox_to_anchor=(.5,-.2))
    plt.tight_layout(); plt.savefig(path,dpi=140,bbox_inches="tight"); plt.close()


def _plot_bar(labels,values,title,xlabel,path):
    plt.figure(figsize=(8.4,4.2)); plt.barh(labels[::-1],list(values)[::-1],color="#1F77B4"); plt.title(title); plt.xlabel(xlabel); plt.tight_layout(); plt.savefig(path,dpi=140); plt.close()


def _historical_metrics(data,ranges):
    result={"Alertas":[],"Riesgo":[],"Fatiga sin cumplir":[],"Sin cinturón":[],"Fumando":[],"Reincidencia %":[]}
    for _,start,end in ranges:
        frame=data[(data["Fecha"]>=start)&(data["Fecha"]<=end)]; risk=operational_risk_summary(frame); rec=recurrence_summary(frame)
        result["Alertas"].append(len(frame)); result["Riesgo"].append(risk["points"]); result["Fatiga sin cumplir"].append(risk["fatigue_no"]); result["Sin cinturón"].append(risk["seatbelt"]); result["Fumando"].append(risk["smoking"]); result["Reincidencia %"].append(round(rec["rate"],1))
    return result


def _methodology_flowables(styles,risk,recurrence):
    rows=[["Factor","Cantidad","Peso","Subtotal"],["Fatiga sin cumplimiento",risk["fatigue_no"],10,risk["fatigue_no"]*10],["Fatiga pendiente",risk["fatigue_pending"],8,risk["fatigue_pending"]*8],["Sin cinturón",risk["seatbelt"],6,risk["seatbelt"]*6],["Conductor fumando",risk["smoking"],6,risk["smoking"]*6],["Fatiga con cumplimiento",risk["fatigue_yes"],3,risk["fatigue_yes"]*3],["TOTAL","","",risk["points"]]]
    explanation=f"<b>Nivel {risk['level']}</b>: {risk['equation']} Regla aplicada: “{risk['rule']}”. Escala: Bajo &lt;10; Medio 10–24; Alto 25–49; Crítico ≥50 o reincidencia de fatiga sin cumplimiento. Cada fatiga sin cumplimiento es individualmente prioridad crítica."
    rec_text=f"Índice de reincidencia: {recurrence['recurrent']} conductores con 2 o más eventos ÷ {recurrence['total']} conductores con eventos = {recurrence['rate']:.1f}%. Se presenta separado del riesgo operacional."
    return [Paragraph("Metodología del riesgo operacional",styles["H2Custom"]),_style_table(Table(rows,colWidths=[7*cm,2.2*cm,2*cm,2.4*cm]),6.2),Spacer(1,5),Paragraph(explanation,styles["BodyCustom"]),Spacer(1,4),Paragraph(rec_text,styles["BodyCustom"])]


def _detail_dataframe(current):
    mapping={"ID":"ID","Fecha":"Fecha","Plataforma":"Plataforma","Conductor":"Conductor","Tracto":"Tracto","Patente":"Patente","Tipo":"Tipo de evento","CumplimientoFatiga":"Cumplimiento fatiga","Fecha/Hora de evento":"Fecha/hora evento","Fecha/Hora de gestion":"Fecha/hora gestión","Tiempo respuesta evento. (No ingresar datos)":"Tiempo respuesta","Velocidad (sólo número)":"Velocidad","Monitor":"Monitor","Observaciones":"Observaciones","Estado":"Estado"}
    cols=[col for col in mapping if col in current.columns]; detail=current[cols].rename(columns=mapping).copy()
    if "Fecha" in detail: detail["Fecha"]=pd.to_datetime(detail["Fecha"]).dt.strftime("%d-%m-%Y")
    return detail.fillna("")


def _write_transportista_excel(path,current,drivers,ranges,metrics,risk,recurrence):
    summary=pd.DataFrame([["Nivel de riesgo operacional",risk["level"]],["Puntos de riesgo",risk["points"]],["Explicación",risk["equation"]],["Regla aplicada",risk["rule"]],["Índice de reincidencia",f"{recurrence['rate']:.1f}%"],["Conductores reincidentes",recurrence["recurrent"]]],columns=["Indicador","Resultado"])
    evolution=pd.DataFrame({"Período":[x[0] for x in ranges],**metrics})
    with pd.ExcelWriter(path,engine="openpyxl") as writer:
        summary.to_excel(writer,sheet_name="Resumen",index=False); drivers.to_excel(writer,sheet_name="Ranking conductores",index=False); evolution.to_excel(writer,sheet_name="Evolución",index=False); _detail_dataframe(current).to_excel(writer,sheet_name="Detalle eventos",index=False)
        for sheet in writer.book.worksheets:
            sheet.freeze_panes="A2"; sheet.auto_filter.ref=sheet.dimensions
            for cell in sheet[1]: cell.font=cell.font.copy(bold=True)
            for column in sheet.columns: sheet.column_dimensions[column[0].column_letter].width=min(max(len(str(cell.value or "")) for cell in column)+2,45)


def _rank_transportistas(current,previous):
    rows=[]; previous_counts=previous.groupby("Transportista").size()
    for transportista,frame in current.groupby("Transportista"):
        risk=operational_risk_summary(frame); rec=recurrence_summary(frame); rows.append({"Transportista":transportista,"Alertas":len(frame),"Anterior":int(previous_counts.get(transportista,0)),"Conductores":frame["Conductor"].nunique(),"Reincidencia":rec["rate"],"Fatiga sin cumplir":risk["fatigue_no"],"Sin cinturón":risk["seatbelt"],"Fumando":risk["smoking"],"Riesgo":risk["points"],"Nivel":risk["level"]})
    rank=pd.DataFrame(rows)
    return rank.sort_values(["Fatiga sin cumplir","Riesgo","Reincidencia","Alertas"],ascending=False) if not rank.empty else rank


def generate_global_report(data,config,output_dir):
    styles=_styles(); current=_period_slice(data,config.period_start,config.period_end); ranges=comparison_ranges(config); labels=[x[0] for x in ranges]; metrics=_historical_metrics(data,ranges); previous=data[(data["Fecha"]>=ranges[-2][1])&(data["Fecha"]<=ranges[-2][2])] if len(ranges)>1 else current.iloc[0:0]
    rank_t=_rank_transportistas(current,previous); drivers=detailed_driver_ranking(current,previous,config.top_drivers,True); risk=operational_risk_summary(current); rec=recurrence_summary(current)
    chart=output_dir/"charts_global"; chart.mkdir(exist_ok=True); p1=chart/"total.png"; p2=chart/"risk.png"; p3=chart/"trans.png"; p4=chart/"drivers.png"
    unit="mensual" if config.report_mode.lower().startswith("mens") else "semanal"; _plot_line(labels,{"Alertas":metrics["Alertas"]},f"Evolución {unit}","Alertas",p1); _plot_line(labels,{"Riesgo":metrics["Riesgo"],"Fatiga sin cumplir":metrics["Fatiga sin cumplir"]},"Evolución del riesgo operacional","Valor",p2); _plot_bar([wrap_label(x,24) for x in rank_t.head(10).Transportista],rank_t.head(10).Riesgo,"Transportistas por riesgo operacional","Puntos",p3); _plot_bar([wrap_label(x,27) for x in drivers.head(10).Conductor],drivers.head(10).Riesgo,"Conductores por riesgo operacional","Puntos",p4)
    pdf=output_dir/f"00_Informe_Global_COPEC_{config.report_mode}_{config.period_start:%d%m}_{config.period_end:%d%m%Y}.pdf"; doc=SimpleDocTemplate(str(pdf),pagesize=A4,rightMargin=1.05*cm,leftMargin=1.05*cm,topMargin=2.25*cm,bottomMargin=1.2*cm)
    kpi=[["Indicador","Resultado"],["Total eventos",len(current)],["Variación",pct_change(len(current),len(previous))],["Transportistas",current.Transportista.nunique()],["Conductores",current.Conductor.nunique()],["Fatiga sin cumplimiento",risk["fatigue_no"]],["Sin cinturón",risk["seatbelt"]],["Conductor fumando",risk["smoking"]],["Reincidencia",f"{rec['rate']:.1f}%"],["Riesgo operacional",f"{risk['level']} ({risk['points']} pts)"]]
    elements=[Paragraph(f"INFORME {config.report_mode.upper()} GLOBAL COPEC",styles["TitleCustom"]),Paragraph(f"Período: {config.period_start:%d-%m-%Y} al {config.period_end:%d-%m-%Y} · Comparación: {config.comparison_periods} períodos",styles["BodyCustom"]),Spacer(1,6),Paragraph("1. Resumen ejecutivo",styles["H2Custom"]),Table([[_style_table(Table(kpi,colWidths=[4.6*cm,2.8*cm]),6.3),Image(str(p1),width=9.4*cm,height=6.1*cm)]],colWidths=[7.6*cm,9.6*cm],style=[("VALIGN",(0,0),(-1,-1),"TOP")]),Spacer(1,5),Image(str(p2),width=17.8*cm,height=6.5*cm),PageBreak()]+_methodology_flowables(styles,risk,rec)+[Spacer(1,7)]
    rows=[["Transportista","Eventos","Fatiga NO","Cint.","Fuma","Reinc. %","Riesgo","Nivel"]]
    for _,r in rank_t.head(15).iterrows(): rows.append([Paragraph(truncate(r.Transportista,31),styles["SmallCustom"]),int(r.Alertas),int(r["Fatiga sin cumplir"]),int(r["Sin cinturón"]),int(r.Fumando),f"{r.Reincidencia:.1f}%",int(r.Riesgo),r.Nivel])
    elements += [Paragraph("2. Transportistas prioritarios",styles["H2Custom"]),Image(str(p3),width=17.8*cm,height=6.7*cm),Spacer(1,4),_style_table(Table(rows,colWidths=[5*cm,1.4*cm,1.5*cm,1.3*cm,1.2*cm,1.5*cm,1.3*cm,1.6*cm]),5.3),PageBreak()]
    drows=[["Conductor","Empresa","Eventos","Reinc.","Fatiga NO","Cint.","Fuma","Riesgo","Nivel"]]
    for _,r in drivers.iterrows(): drows.append([Paragraph(truncate(r.Conductor,27),styles["SmallCustom"]),Paragraph(truncate(r.Transportista,22),styles["SmallCustom"]),int(r.Alertas),int(r.Reincidencias),int(r["Fatiga sin cumplir"]),int(r["Sin cinturón"]),int(r.Fumando),int(r.Riesgo),r.Nivel])
    elements += [Paragraph("3. Ranking operativo de conductores",styles["H2Custom"]),Image(str(p4),width=17.8*cm,height=6.8*cm),Spacer(1,4),_style_table(Table(drows,colWidths=[4*cm,3.3*cm,1.1*cm,1.1*cm,1.2*cm,1*cm,1*cm,1.1*cm,1.4*cm]),5.0),Spacer(1,5),Paragraph("Orden: fatiga sin cumplimiento, reincidencia, sin cinturón/conductor fumando y volumen total.",styles["BodyCustom"])]
    doc.build(elements,onFirstPage=_page_branding,onLaterPages=_page_branding); return pdf


def generate_transportista_report(data,transportista,config,output_dir,index):
    styles=_styles(); all_t=data[data.Transportista==transportista].copy(); current=_period_slice(all_t,config.period_start,config.period_end); ranges=comparison_ranges(config); labels=[x[0] for x in ranges]; metrics=_historical_metrics(all_t,ranges); previous=all_t[(all_t.Fecha>=ranges[-2][1])&(all_t.Fecha<=ranges[-2][2])] if len(ranges)>1 else current.iloc[0:0]
    drivers=detailed_driver_ranking(current,previous,config.top_drivers,False); rec=recurrence_summary(current); risk=operational_risk_summary(current)
    chart=output_dir/f"charts_{index:02d}"; chart.mkdir(exist_ok=True); p1=chart/"evol.png"; p2=chart/"risk.png"; p3=chart/"drivers.png"; unit="mensual" if config.report_mode.lower().startswith("mens") else "semanal"
    _plot_line(labels,{"Alertas":metrics["Alertas"]},f"Evolución {unit}","Alertas",p1); _plot_line(labels,{"Riesgo":metrics["Riesgo"],"Fatiga sin cumplir":metrics["Fatiga sin cumplir"],"Sin cinturón":metrics["Sin cinturón"],"Fumando":metrics["Fumando"]},"Riesgo operacional por período","Valor",p2); _plot_bar([wrap_label(x,26) for x in drivers.Conductor],drivers.Riesgo,"Ranking operativo de conductores","Puntos",p3)
    pdf=output_dir/f"{index:02d}_Informe_{config.report_mode}_{safe_filename(transportista)}_{config.period_start:%d%m}_{config.period_end:%d%m%Y}.pdf"; doc=SimpleDocTemplate(str(pdf),pagesize=A4,rightMargin=1*cm,leftMargin=1*cm,topMargin=2.25*cm,bottomMargin=1.2*cm)
    kpi=[["Indicador","Resultado"],["Eventos",len(current)],["Período anterior",len(previous)],["Variación",pct_change(len(current),len(previous))],["Conductores",current.Conductor.nunique()],["Equipos",current.Tracto.nunique()],["Fatiga sin cumplimiento",risk["fatigue_no"]],["Sin cinturón",risk["seatbelt"]],["Conductor fumando",risk["smoking"]],["Reincidencia",f"{rec['rate']:.1f}%"],["Riesgo operacional",f"{risk['level']} ({risk['points']} pts)"]]
    priority=[["Prioridad","Tipo","Eventos","Conductores","Reincidentes"],["CRÍTICA","Fatiga sin cumplimiento",risk["fatigue_no"],current[(current.Tipo=="Fatiga")&(current.CumplimientoFatiga=="NO")].Conductor.nunique(),risk["repeated_fatigue_no_drivers"]],["ALTA","Sin cinturón",risk["seatbelt"],current[current.Tipo=="Sin cinturón"].Conductor.nunique(),int((current[current.Tipo=="Sin cinturón"].groupby("Conductor").size()>=2).sum())],["ALTA","Conductor fumando",risk["smoking"],current[current.Tipo=="Conductor fumando"].Conductor.nunique(),int((current[current.Tipo=="Conductor fumando"].groupby("Conductor").size()>=2).sum())],["MEDIA","Fatiga pendiente",risk["fatigue_pending"],current[(current.Tipo=="Fatiga")&(current.CumplimientoFatiga=="PENDIENTE")].Conductor.nunique(),0]]
    elements=[Paragraph(f"INFORME {config.report_mode.upper()} DE ALERTAS",styles["TitleCustom"]),Paragraph(f"Transportista: {transportista}<br/>Período: {config.period_start:%d-%m-%Y} al {config.period_end:%d-%m-%Y} · Comparación: {config.comparison_periods} períodos",styles["BodyCustom"]),Spacer(1,6),Paragraph("1. Resumen ejecutivo",styles["H2Custom"]),Table([[_style_table(Table(kpi,colWidths=[4.7*cm,2.7*cm]),6.1),Image(str(p1),width=9.4*cm,height=6.1*cm)]],colWidths=[7.7*cm,9.5*cm],style=[("VALIGN",(0,0),(-1,-1),"TOP")]),Spacer(1,4),_style_table(Table(priority,colWidths=[1.7*cm,5.1*cm,2*cm,2.2*cm,2.2*cm]),5.8),Spacer(1,4),Image(str(p2),width=17.8*cm,height=6.0*cm),PageBreak()]+_methodology_flowables(styles,risk,rec)+[Spacer(1,7),Paragraph("2. Ranking operativo de conductores",styles["H2Custom"]),Image(str(p3),width=17.8*cm,height=6.5*cm),Spacer(1,4)]
    rows=[["Conductor","Eventos","Reinc.","Fatiga NO","Pend.","Cint.","Fuma","Riesgo","Nivel"]]
    for _,r in drivers.iterrows(): rows.append([Paragraph(truncate(r.Conductor,34),styles["SmallCustom"]),int(r.Alertas),int(r.Reincidencias),int(r["Fatiga sin cumplir"]),int(r["Fatiga pendiente"]),int(r["Sin cinturón"]),int(r.Fumando),int(r.Riesgo),r.Nivel])
    elements += [_style_table(Table(rows,colWidths=[5.2*cm,1.1*cm,1.1*cm,1.2*cm,1*cm,1*cm,1*cm,1.1*cm,1.4*cm]),5.0),Spacer(1,4),Paragraph("Orden: fatiga sin cumplimiento, reincidencia, sin cinturón/conductor fumando y total de eventos.",styles["BodyCustom"]),PageBreak()]
    actions=[["Prioridad","Hallazgo","Acción recomendada"],["CRÍTICA",f"{risk['fatigue_no']} fatigas sin cumplimiento","Revisar de inmediato descansos, llamados y trazabilidad; definir responsable y fecha de cierre."],["ALTA",f"{risk['seatbelt']} sin cinturón y {risk['smoking']} fumando","Retroalimentación individual, verificación de reincidencia y control en el período siguiente."],["MEDIA",f"{risk['fatigue_pending']} fatigas pendientes","Completar el registro de cumplimiento y cerrar la gestión pendiente."],["SEGUIMIENTO",f"Índice de reincidencia {rec['rate']:.1f}%","Monitorear conductores con dos o más eventos y documentar acciones correctivas."]]
    elements += [Paragraph("3. Plan de acción",styles["H2Custom"]),_style_table(Table([[Paragraph(str(c),styles["SmallCustom"]) for c in row] for row in actions],colWidths=[2*cm,6*cm,9*cm]),5.8),PageBreak(),Paragraph("4. Detalle completo de eventos",styles["H2Custom"])]
    detail=_detail_dataframe(current); detail_cols=[c for c in ["ID","Fecha","Plataforma","Conductor","Tracto","Patente","Tipo de evento","Cumplimiento fatiga","Fecha/hora evento","Velocidad","Monitor","Estado"] if c in detail.columns]; drows=[detail_cols]
    for _,r in detail[detail_cols].iterrows(): drows.append([Paragraph(truncate(r[col],24),styles["TinyCustom"]) for col in detail_cols])
    widths=[1.4,1.5,1.4,3.3,1.2,1.4,2.5,1.7,2.1,1,1.6,1.2][:len(detail_cols)]; elements += [_style_table(LongTable(drows,colWidths=[x*cm for x in widths],repeatRows=1),4.7)]
    doc.build(elements,onFirstPage=_page_branding,onLaterPages=_page_branding)
    excel=output_dir/f"{index:02d}_Detalle_{safe_filename(transportista)}_{config.period_start:%d%m}_{config.period_end:%d%m%Y}.xlsx"; _write_transportista_excel(excel,current,drivers,ranges,metrics,risk,rec)
    return pdf,excel


def generate_reports(data,config):
    data=filter_real_transportistas(data); cutoff=pd.Timestamp(datetime.combine(config.period_end,datetime.max.time())); data=data[data["Fecha"]<=cutoff].copy(); current=_period_slice(data,config.period_start,config.period_end)
    if current.empty: raise ValueError("No existen alertas dentro del período seleccionado.")
    with tempfile.TemporaryDirectory(prefix="copec_reports_") as tmp:
        out=Path(tmp); global_bytes=None; files=[]; global_path=None
        if config.include_global: global_path=generate_global_report(data,config,out); global_bytes=global_path.read_bytes()
        transportistas=sorted(current.Transportista.dropna().unique()) if config.include_transportistas else []
        for index,transportista in enumerate(transportistas,1): files.extend(generate_transportista_report(data,transportista,config,out,index))
        buffer=io.BytesIO()
        with zipfile.ZipFile(buffer,"w",zipfile.ZIP_DEFLATED) as archive:
            if global_path: archive.write(global_path,arcname=global_path.name)
            for path in files: archive.write(path,arcname=path.name)
        zip_bytes=buffer.getvalue()
    risk=operational_risk_summary(current); rec=recurrence_summary(current)
    return global_bytes,zip_bytes,{"total_alertas":len(current),"transportistas":len(transportistas),"conductores":current.Conductor.nunique(),"equipos":current.Tracto.nunique(),"fatiga":int((current.Tipo=="Fatiga").sum()),"riesgo":risk["level"],"reincidencia":rec["rate"],"modo":config.report_mode}
