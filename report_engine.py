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

from rco_sections import identified_data, driver_count, driver_keys, rco_summary, fmt_pct, executive_pages, rco_appendix, body

from alert_aliases import SENSOR_BLOCKED, SENSOR_MISALIGNED, normalize_alert_name

RISK_WEIGHTS = {"Fatiga sin cumplimiento": 10, "Fatiga pendiente": 8, "Sin cinturón": 6, "Conductor fumando": 6, "Fatiga con cumplimiento": 3}
GENERIC_TRANSPORTISTAS = {"COPEC", "NO ES COPEC", "PLANTA", "DESCONOCIDO", "OWL", "GPS", "PRUEBA", "SIN TRANSPORTISTA"}
REPORT_ENGINE_VERSION = "2026.09.14-rco-1"
REPORT_LOGO = Path(__file__).resolve().parent / "copec.png"


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
    # Las categorías genéricas se excluyen sólo por coincidencia exacta. Una
    # empresa real puede contener palabras como COPEC, GPS o PLANTA en su nombre.
    normalized = data["Transportista"].map(norm_text)
    return data[~normalized.isin(GENERIC_TRANSPORTISTAS)].copy()


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
    df = identified_data(df)
    counts = df.groupby(driver_keys(df)).size() if not df.empty else pd.Series(dtype=int); total = len(counts); recurrent = int((counts>=2).sum())
    return {"one":int((counts==1).sum()),"two":int((counts==2).sum()),"three_plus":int((counts>=3).sum()),"recurrent":recurrent,"total":total,"rate":recurrent/total*100 if total else float("nan")}


def operational_risk_summary(df):
    fatigue = df[df["Tipo"].eq("Fatiga")]
    no = int((fatigue["CumplimientoFatiga"]=="NO").sum()); pending = int((fatigue["CumplimientoFatiga"]=="PENDIENTE").sum()); yes = int((fatigue["CumplimientoFatiga"]=="SI").sum())
    belt = int((df["Tipo"]=="Sin cinturón").sum()); smoking = int((df["Tipo"]=="Conductor fumando").sum())
    points = no*10+pending*8+belt*6+smoking*6+yes*3
    repeated = int((identified_data(fatigue[fatigue["CumplimientoFatiga"]=="NO"]).groupby(driver_keys(df)).size()>=2).sum())
    if points >= 50 or repeated: level, rule = "CRÍTICO", ("50 puntos o más" if points>=50 else "reincidencia de fatiga sin cumplimiento")
    elif points >= 25: level, rule = "ALTO", "entre 25 y 49 puntos"
    elif points >= 10: level, rule = "MEDIO", "entre 10 y 24 puntos"
    else: level, rule = "BAJO", "menos de 10 puntos"
    equation = f"{no}×10 (fatiga sin cumplimiento) + {pending}×8 (fatiga pendiente) + {belt}×6 (sin cinturón) + {smoking}×6 (conductor fumando) + {yes}×3 (fatiga con cumplimiento) = {points} puntos."
    return {"fatigue_no":no,"fatigue_pending":pending,"fatigue_yes":yes,"seatbelt":belt,"smoking":smoking,"points":points,"level":level,"rule":rule,"repeated_fatigue_no_drivers":repeated,"equation":equation}


def detailed_driver_ranking(current, previous, topn=10, include_transportista=False):
    current = identified_data(current); previous = identified_data(previous)
    group = ["Conductor","Transportista"] if include_transportista else ["Conductor"]; rows=[]
    if current.empty: return pd.DataFrame(columns=group+["Alertas","Reincidencias","Fatiga sin cumplir","Fatiga pendiente","Sin cinturón","Fumando","Riesgo","Nivel","Anterior","Variación"])
    previous_counts = previous.groupby(group).size()
    for keys, frame in current.groupby(group, dropna=False):
        keys = keys if isinstance(keys,tuple) else (keys,); risk=operational_risk_summary(frame); alerts=len(frame)
        row={group[i]:keys[i] for i in range(len(group))}; row.update({"Alertas":alerts,"Reincidencias":max(alerts-1,0),"Fatiga sin cumplir":risk["fatigue_no"],"Fatiga pendiente":risk["fatigue_pending"],"Sin cinturón":risk["seatbelt"],"Fumando":risk["smoking"],"Riesgo":risk["points"],"Nivel":risk["level"],"Anterior":int(previous_counts.get(keys if include_transportista else keys[0],0))})
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
        scale = logo_width / 679
        x = 1.05 * cm; y = page_height - 0.65 * cm - 143 * scale
        canvas.saveState()
        clip = canvas.beginPath(); clip.rect(x, y, logo_width, 143 * scale)
        canvas.clipPath(clip, stroke=0)
        canvas.drawImage(str(REPORT_LOGO), x-94*scale, y-253*scale,
                         width=866*scale, height=650*scale)
        canvas.restoreState()
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
    rec_text=f"Índice de reincidencia: {recurrence['recurrent']} conductores con 2 o más eventos ÷ {recurrence['total']} conductores con eventos = {fmt_pct(recurrence['rate'])}. Se presenta separado del riesgo operacional."
    return [Paragraph("Metodología del riesgo operacional",styles["H2Custom"]),_style_table(Table(rows,colWidths=[7*cm,2.2*cm,2*cm,2.4*cm]),6.2),Spacer(1,5),Paragraph(explanation,styles["BodyCustom"]),Spacer(1,4),Paragraph(rec_text,styles["BodyCustom"])]


def _detail_dataframe(current):
    mapping={"ID":"ID","Fecha":"Fecha","Plataforma":"Plataforma","Conductor":"Conductor","Tracto":"Tracto","Patente":"Patente","Tipo":"Tipo de evento","CumplimientoFatiga":"Cumplimiento fatiga"}
    cols=[col for col in mapping if col in current.columns]; detail=current[cols].rename(columns=mapping).copy()
    if "Fecha" in detail: detail["Fecha"]=pd.to_datetime(detail["Fecha"]).dt.strftime("%d-%m-%Y")
    return detail.fillna("")


def _rank_transportistas(current,previous):
    rows=[]; previous_counts=previous.groupby("Transportista").size()
    for transportista,frame in current.groupby("Transportista"):
        risk=operational_risk_summary(frame); rec=recurrence_summary(frame); rows.append({"Transportista":transportista,"Alertas":len(frame),"Anterior":int(previous_counts.get(transportista,0)),"Conductores":driver_count(frame),"Reincidencia":rec["rate"],"Fatiga sin cumplir":risk["fatigue_no"],"Sin cinturón":risk["seatbelt"],"Fumando":risk["smoking"],"Riesgo":risk["points"],"Nivel":risk["level"]})
    rank=pd.DataFrame(rows)
    return rank.sort_values(["Fatiga sin cumplir","Riesgo","Reincidencia","Alertas"],ascending=False) if not rank.empty else rank


def _alert_history(data, ranges, limit=6):
    current = data[(data["Fecha"] >= ranges[-1][1]) & (data["Fecha"] <= ranges[-1][2])]
    ordered = current["Tipo"].value_counts().index.tolist()
    preferred = ["Sin cinturón", "Fatiga", "Uso celular", "Conductor fumando", SENSOR_MISALIGNED, SENSOR_BLOCKED]
    selected = []
    for name in preferred + ordered:
        if name in ordered and name not in selected:
            selected.append(name)
        if len(selected) == limit:
            break
    values = {name: [] for name in selected}
    for _, start, end in ranges:
        frame = data[(data["Fecha"] >= start) & (data["Fecha"] <= end)]
        counts = frame["Tipo"].value_counts()
        for name in selected:
            values[name].append(int(counts.get(name, 0)))
    return values


def _plot_alert_small_multiples(labels, series, path, period_label="semanal"):
    count = max(len(series), 1); cols = 3; rows = int(np.ceil(count / cols))
    fig, axes = plt.subplots(rows, cols, figsize=(14, 3.2 * rows), squeeze=False, facecolor="white")
    colors_list = ["#E31B16", "#F28E2B", "#2478B8", "#7B4AB4", "#169C82", "#4E79A7"]
    for ax in axes.flat[count:]: ax.axis("off")
    for ax, (name, values), color in zip(axes.flat, series.items(), colors_list):
        x = range(len(labels)); ax.plot(x, values, marker="o", linewidth=2.8, markersize=5, color=color)
        ax.fill_between(x, values, [min(values) * .92 if values else 0] * len(values), alpha=.08, color=color)
        previous = values[-2] if len(values) > 1 else 0; change = pct_value(values[-1], previous)
        ax.set_title(name, loc="left", fontsize=12, fontweight="bold", color="#222222")
        ax.text(.99, .94, f"{values[-1]}  {change:+.1f}%", transform=ax.transAxes, ha="right", va="top", fontsize=11, fontweight="bold", color="#C52620" if change > 0 else "#16835E")
        ax.set_xticks(list(x), labels, rotation=18, ha="right", fontsize=8); ax.tick_params(axis="y", labelsize=8, length=0); ax.tick_params(axis="x", length=0)
        ax.grid(axis="y", alpha=.18); ax.spines[["top", "right", "left"]].set_visible(False)
    fig.suptitle(f"Evolución {period_label} por tipo de alerta", fontsize=18, fontweight="bold", color="#1F4E79", y=.98)
    plt.tight_layout(rect=[0, .01, 1, .93], h_pad=1.7, w_pad=1.2); plt.savefig(path, dpi=180, bbox_inches="tight"); plt.close()


def _trend_rows(current, previous):
    now = current["Tipo"].value_counts(); before = previous["Tipo"].value_counts(); total = max(len(current), 1); rows=[]
    for name, value in now.items():
        old = int(before.get(name, 0)); change = pct_value(int(value), old)
        trend = "AUMENTA" if change > 5 else "DISMINUYE" if change < -5 else "ESTABLE"
        rows.append([name, int(value), old, f"{change:+.1f}%", f"{value/total*100:.1f}%", trend])
    return [["Tipo de alerta","Actual","Anterior","Variación","Participación","Tendencia"]] + rows


def _executive_findings(current, previous, risk, recurrence, subject="COPEC"):
    findings=[]; total_change=pct_value(len(current),len(previous)); direction="aumentaron" if total_change>0 else "disminuyeron" if total_change<0 else "se mantuvieron"
    findings.append(f"Los eventos de {subject} {direction} {abs(total_change):.1f}%: {len(current)} actuales frente a {len(previous)} del período anterior.")
    now=current["Tipo"].value_counts(); before=previous["Tipo"].value_counts(); candidates=[]
    for name,value in now.items(): candidates.append((int(value)-int(before.get(name,0)),name,int(value),int(before.get(name,0))))
    if candidates:
        delta,name,value,old=max(candidates)
        findings.append(f"La mayor presión proviene de {name}: {value} eventos, {delta:+d} respecto del período anterior.")
    findings.append(f"Existen {risk['fatigue_no']} fatigas sin cumplimiento; {risk['repeated_fatigue_no_drivers']} conductores repiten este incumplimiento.")
    findings.append(f"La reincidencia alcanza {fmt_pct(recurrence['rate'])}: {recurrence['recurrent']} de {recurrence['total']} conductores registran dos o más eventos.")
    return findings


def _findings_box(styles, findings, width=9.5*cm):
    rows=[[Paragraph(f"<b>{i}.</b> {text}",styles["BodyCustom"])] for i,text in enumerate(findings,1)]
    table=Table(rows,colWidths=[width]); table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#F2F7FB")),("BOX",(0,0),(-1,-1),.7,colors.HexColor("#8BAAC5")),("INNERGRID",(0,0),(-1,-1),.2,colors.HexColor("#DCE7EF")),("LEFTPADDING",(0,0),(-1,-1),7),("RIGHTPADDING",(0,0),(-1,-1),7),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)])); return table


def _action_rows(risk, rec, owner="Operaciones / Transportista"):
    return [["Prioridad","Hallazgo","Acción recomendada","Responsable","Plazo"],
        ["INMEDIATA",f"{risk['fatigue_no']} fatigas sin cumplimiento","Revisar descanso, contacto y trazabilidad; documentar cierre.",owner,"24 h"],
        ["ALTA",f"{risk['seatbelt']} eventos sin cinturón","Intervenir reincidentes y verificar resultado en el siguiente período.","Prevención","7 días"],
        ["ALTA",f"Reincidencia {fmt_pct(rec['rate'])}","Asignar seguimiento individual a conductores con dos o más eventos.","Jefatura operativa","7 días"],
        ["TÉCNICA",f"{risk['fatigue_pending']} fatigas pendientes","Completar clasificación y cerrar gestiones pendientes.","Torre de Control","48 h"]]


def sensor_alert_summary(current, previous):
    sensor_types=[SENSOR_BLOCKED,SENSOR_MISALIGNED]
    now=current[current["Tipo"].isin(sensor_types)].copy(); old=previous[previous["Tipo"].isin(sensor_types)].copy()
    unit="Tracto" if "Tracto" in current.columns else "Patente"
    unit_counts=now.groupby(unit).size() if not now.empty else pd.Series(dtype=int)
    recurrent=int((unit_counts>=2).sum()); affected=int(unit_counts.size)
    both=int((now.groupby(unit)["Tipo"].nunique()>=2).sum()) if not now.empty else 0
    rows=[]
    for alert_type in sensor_types:
        frame=now[now["Tipo"].eq(alert_type)]; before=old[old["Tipo"].eq(alert_type)]; counts=frame.groupby(unit).size() if not frame.empty else pd.Series(dtype=int)
        rows.append({"Tipo":alert_type,"Actual":len(frame),"Anterior":len(before),"Variación":pct_value(len(frame),len(before)),"Tractos":int(counts.size),"Reincidentes":int((counts>=2).sum()),"Participación":len(frame)/len(now)*100 if len(now) else 0.0})
    variation=pct_value(len(now),len(old)); high=variation>20 or recurrent>=5 or both>0
    priority="ALTA" if high else "MEDIA" if len(now) or recurrent else "BAJA"
    return {"rows":rows,"total":len(now),"previous":len(old),"variation":variation,"affected":affected,"recurrent":recurrent,"both":both,"priority":priority,"unit":unit}


def _sensor_summary_table(styles, summary):
    rows=[["Tipo de evento","Actual","Anterior","Variación","Tractos","Reincidentes","Participación","Tendencia"]]
    for item in summary["rows"]:
        trend="AUMENTA" if item["Variación"]>5 else "DISMINUYE" if item["Variación"]<-5 else "ESTABLE"
        rows.append([item["Tipo"],item["Actual"],item["Anterior"],f"{item['Variación']:+.1f}%",item["Tractos"],item["Reincidentes"],f"{item['Participación']:.1f}%",trend])
    rows.append(["TOTAL SENSORES",summary["total"],summary["previous"],f"{summary['variation']:+.1f}%",summary["affected"],summary["recurrent"],"100.0%" if summary["total"] else "0.0%",summary["priority"]])
    return _style_table(Table(rows,colWidths=[4.1*cm,1.4*cm,1.4*cm,1.7*cm,1.5*cm,1.8*cm,1.8*cm,2.1*cm]),5.4)


def _sensor_cards(styles, summary):
    blocked=next(x for x in summary["rows"] if x["Tipo"]==SENSOR_BLOCKED); misaligned=next(x for x in summary["rows"] if x["Tipo"]==SENSOR_MISALIGNED)
    data=[["Eventos sensores","Tractos afectados",SENSOR_BLOCKED,SENSOR_MISALIGNED],[f"{summary['total']} ({summary['variation']:+.1f}%)",f"{summary['affected']} · {summary['recurrent']} reinc.",f"{blocked['Actual']} ({blocked['Variación']:+.1f}%)",f"{misaligned['Actual']} ({misaligned['Variación']:+.1f}%)"]]
    table=Table(data,colWidths=[4.25*cm]*4,rowHeights=[.65*cm,1.05*cm]); table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#F2F7FB")),("TEXTCOLOR",(0,0),(-1,0),colors.HexColor("#333333")),("TEXTCOLOR",(0,1),(-1,1),colors.HexColor("#1F4E79")),("FONTNAME",(0,0),(-1,0),"Helvetica"),("FONTNAME",(0,1),(-1,1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,0),7),("FONTSIZE",(0,1),(-1,1),11),("BOX",(0,0),(-1,-1),.5,colors.HexColor("#BCD0E0")),("INNERGRID",(0,0),(-1,-1),.3,colors.HexColor("#D7E3EC")),("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
    return table


def _sensor_transportista_rows(current, previous):
    rows=[["Transportista","Tapado","Desalineado","Total","Tractos","Reincidentes","Variación","Prioridad"]]
    for transportista,frame in current.groupby("Transportista"):
        old=previous[previous["Transportista"].eq(transportista)]; summary=sensor_alert_summary(frame,old)
        if not summary["total"]: continue
        blocked=next(x for x in summary["rows"] if x["Tipo"]==SENSOR_BLOCKED); misaligned=next(x for x in summary["rows"] if x["Tipo"]==SENSOR_MISALIGNED)
        rows.append([transportista,blocked["Actual"],misaligned["Actual"],summary["total"],summary["affected"],summary["recurrent"],f"{summary['variation']:+.1f}%",summary["priority"]])
    return rows


def _sensor_equipment_rows(current):
    sensors=current[current["Tipo"].isin([SENSOR_BLOCKED,SENSOR_MISALIGNED])].copy(); rows=[["Tracto","Patente","Tapado","Desalineado","Total","Último evento","Prioridad"]]
    if sensors.empty: return rows
    for tracto,frame in sensors.groupby("Tracto"):
        blocked=int((frame["Tipo"]==SENSOR_BLOCKED).sum()); misaligned=int((frame["Tipo"]==SENSOR_MISALIGNED).sum()); total=len(frame)
        priority="ALTA" if total>=3 or (blocked and misaligned) else "MEDIA" if total>=2 else "SEGUIMIENTO"
        patente=frame["Patente"].dropna().astype(str).iloc[0] if frame["Patente"].notna().any() else "S/I"
        rows.append([tracto,patente,blocked,misaligned,total,pd.to_datetime(frame["Fecha"]).max().strftime("%d-%m-%Y"),priority])
    return [rows[0]]+sorted(rows[1:],key=lambda row:row[4],reverse=True)[:15]


def generate_global_report(data,config,output_dir):
    styles=_styles(); current=_period_slice(data,config.period_start,config.period_end); ranges=comparison_ranges(config); labels=[x[0] for x in ranges]; previous=data[(data["Fecha"]>=ranges[-2][1])&(data["Fecha"]<=ranges[-2][2])] if len(ranges)>1 else current.iloc[0:0]
    risk=operational_risk_summary(current); rec=recurrence_summary(current); rank_t=_rank_transportistas(current,previous); drivers=detailed_driver_ranking(current,previous,max(config.top_drivers,15),True)
    chart=output_dir/"charts_global"; chart.mkdir(exist_ok=True); alert_chart=chart/"alert_types.png"; trans_chart=chart/"transportistas.png"; driver_chart=chart/"drivers.png"
    period_label="mensual" if config.report_mode.lower().startswith("mens") else "semanal"; _plot_alert_small_multiples(labels,_alert_history(data,ranges),alert_chart,period_label); _plot_bar([wrap_label(x,24) for x in rank_t.head(10).Transportista],rank_t.head(10).Riesgo,"Transportistas por riesgo operacional","Puntos",trans_chart); _plot_bar([wrap_label(x,27) for x in drivers.head(10).Conductor],drivers.head(10).Riesgo,"Conductores por riesgo operacional","Puntos",driver_chart)
    pdf=output_dir/f"00_Informe_Global_COPEC_{config.report_mode}_{config.period_start:%d%m}_{config.period_end:%d%m%Y}.pdf"; doc=SimpleDocTemplate(str(pdf),pagesize=A4,rightMargin=1.05*cm,leftMargin=1.05*cm,topMargin=2.25*cm,bottomMargin=1.2*cm)
    kpi=[["Indicador","Resultado"],["Eventos",len(current)],["Variación",pct_change(len(current),len(previous))],["Transportistas",current.Transportista.nunique()],["Conductores identificados",driver_count(current)],["Fatiga sin cumplimiento",risk["fatigue_no"]],["Reincidencia",f"{fmt_pct(rec['rate'])}"],["Riesgo operacional",f"{risk['level']} ({risk['points']} pts)"]]
    findings=_executive_findings(current,previous,risk,rec); trend=_trend_rows(current,previous); sensor_summary=sensor_alert_summary(current,previous); sensor_rank=_sensor_transportista_rows(current,previous)
    elements=[Paragraph(f"INFORME {config.report_mode.upper()} GLOBAL COPEC",styles["TitleCustom"]),Paragraph(f"Período: {config.period_start:%d-%m-%Y} al {config.period_end:%d-%m-%Y} · Comparación: {config.comparison_periods} períodos",styles["BodyCustom"]),Spacer(1,5),Paragraph("1. Resumen ejecutivo y decisiones",styles["H2Custom"]),Table([[_style_table(Table(kpi,colWidths=[4.6*cm,2.8*cm]),6.1),_findings_box(styles,findings)]],colWidths=[7.5*cm,9.7*cm],style=[("VALIGN",(0,0),(-1,-1),"TOP")]),Spacer(1,5),Image(str(alert_chart),width=17.6*cm,height=8.0*cm),PageBreak(),Paragraph("2. Evolución y composición de las alertas",styles["H2Custom"]),_style_table(Table(trend[:13],colWidths=[5.2*cm,2*cm,2*cm,2.2*cm,2.2*cm,2.3*cm]),5.8),Spacer(1,8)]+_methodology_flowables(styles,risk,rec)+[PageBreak(),Paragraph("3. Análisis técnico de alertas de sensores",styles["H2Custom"]),_sensor_cards(styles,sensor_summary),Spacer(1,7),_sensor_summary_table(styles,sensor_summary),Spacer(1,10),Paragraph("Transportistas con mayor incidencia técnica",styles["H2Custom"]),_style_table(Table([[Paragraph(truncate(c,30),styles["SmallCustom"]) if i==0 and r>0 else c for i,c in enumerate(row)] for r,row in enumerate(sensor_rank)],colWidths=[5.1*cm,1.3*cm,1.8*cm,1.3*cm,1.4*cm,1.7*cm,1.7*cm,1.8*cm]),5.0),Spacer(1,7),Paragraph("Prioridad alta: aumento superior al 20%, cinco o más tractos reincidentes o ambos tipos de alerta en el mismo equipo.",styles["BodyCustom"]),PageBreak(),Paragraph("4. Transportistas prioritarios",styles["H2Custom"]),Image(str(trans_chart),width=17.5*cm,height=6.8*cm),Spacer(1,4)]
    rows=[["Transportista","Eventos","Fatiga NO","Cint.","Fuma","Reinc. %","Riesgo","Motivo"]]
    for _,r in rank_t.head(15).iterrows():
        reason="Fatiga" if r["Fatiga sin cumplir"] else "Reincidencia" if r.Reincidencia>=50 else "Cinturón" if r["Sin cinturón"] else "Volumen"
        rows.append([Paragraph(truncate(r.Transportista,30),styles["SmallCustom"]),int(r.Alertas),int(r["Fatiga sin cumplir"]),int(r["Sin cinturón"]),int(r.Fumando),f"{fmt_pct(r.Reincidencia)}",int(r.Riesgo),reason])
    elements += [_style_table(Table(rows,colWidths=[5.1*cm,1.3*cm,1.5*cm,1.2*cm,1.1*cm,1.5*cm,1.3*cm,2*cm]),5.1),PageBreak(),Paragraph("5. Conductores y concentración del riesgo",styles["H2Custom"]),Image(str(driver_chart),width=17.5*cm,height=6.5*cm),Spacer(1,4)]
    drows=[["Conductor","Empresa","Eventos","Reinc.","Fatiga NO","Cint.","Fuma","Riesgo","Nivel"]]
    for _,r in drivers.head(config.top_drivers).iterrows(): drows.append([Paragraph(truncate(r.Conductor,25),styles["SmallCustom"]),Paragraph(truncate(r.Transportista,20),styles["SmallCustom"]),int(r.Alertas),int(r.Reincidencias),int(r["Fatiga sin cumplir"]),int(r["Sin cinturón"]),int(r.Fumando),int(r.Riesgo),r.Nivel])
    top_risk=drivers.head(10).Riesgo.sum() if not drivers.empty else 0; concentration=top_risk/risk["points"]*100 if risk["points"] else 0
    elements += [_style_table(Table(drows,colWidths=[3.8*cm,3.1*cm,1.1*cm,1.1*cm,1.2*cm,1*cm,1*cm,1.1*cm,1.4*cm]),4.9),Spacer(1,6),Paragraph(f"Los 10 conductores principales concentran {concentration:.1f}% de los puntos de riesgo del período.",styles["BodyCustom"]),PageBreak(),Paragraph("6. Plan de acción y seguimiento",styles["H2Custom"]),_style_table(Table([[Paragraph(str(c),styles["SmallCustom"]) for c in row] for row in _action_rows(risk,rec)],colWidths=[1.8*cm,4.4*cm,6.4*cm,3.1*cm,1.3*cm]),5.5),Spacer(1,8),Paragraph("Criterio de seguimiento",styles["H2Custom"]),Paragraph("En el próximo informe se debe verificar si disminuyeron las alertas intervenidas, si los conductores reincidentes repiten eventos y si las fatigas pendientes fueron cerradas. Las acciones sin mejora deben escalarse al responsable operacional.",styles["BodyCustom"])]
    elements = executive_pages(current, previous, config, risk=risk) + elements + rco_appendix(current, previous, global_report=True)
    doc.build(elements,onFirstPage=_page_branding,onLaterPages=_page_branding); return pdf


def generate_transportista_report(data,transportista,config,output_dir,index):
    styles=_styles(); all_t=data[data.Transportista==transportista].copy(); current=_period_slice(all_t,config.period_start,config.period_end); ranges=comparison_ranges(config); labels=[x[0] for x in ranges]; previous=all_t[(all_t.Fecha>=ranges[-2][1])&(all_t.Fecha<=ranges[-2][2])] if len(ranges)>1 else current.iloc[0:0]
    drivers=detailed_driver_ranking(current,previous,config.top_drivers,False); rec=recurrence_summary(current); risk=operational_risk_summary(current)
    chart=output_dir/f"charts_{index:02d}"; chart.mkdir(exist_ok=True); p1=chart/"alert_types.png"; p3=chart/"drivers.png"
    period_label="mensual" if config.report_mode.lower().startswith("mens") else "semanal"; _plot_alert_small_multiples(labels,_alert_history(all_t,ranges),p1,period_label); _plot_bar([wrap_label(x,26) for x in drivers.Conductor],drivers.Riesgo,"Conductores por riesgo operacional","Puntos",p3)
    pdf=output_dir/f"{index:02d}_Informe_{config.report_mode}_{safe_filename(transportista)}_{config.period_start:%d%m}_{config.period_end:%d%m%Y}.pdf"; doc=SimpleDocTemplate(str(pdf),pagesize=A4,rightMargin=1*cm,leftMargin=1*cm,topMargin=2.25*cm,bottomMargin=1.2*cm)
    kpi=[["Indicador","Resultado"],["Eventos",len(current)],["Período anterior",len(previous)],["Variación",pct_change(len(current),len(previous))],["Conductores identificados",driver_count(current)],["Equipos",current.Tracto.nunique()],["Fatiga sin cumplimiento",risk["fatigue_no"]],["Sin cinturón",risk["seatbelt"]],["Conductor fumando",risk["smoking"]],["Reincidencia",f"{fmt_pct(rec['rate'])}"],["Riesgo operacional",f"{risk['level']} ({risk['points']} pts)"]]
    findings=_executive_findings(current,previous,risk,rec,transportista); trend=_trend_rows(current,previous); sensor_summary=sensor_alert_summary(current,previous); equipment_rows=_sensor_equipment_rows(current)
    elements=[Paragraph(f"INFORME {config.report_mode.upper()} DE ALERTAS",styles["TitleCustom"]),Paragraph(f"Transportista: {transportista}<br/>Período: {config.period_start:%d-%m-%Y} al {config.period_end:%d-%m-%Y} · Comparación: {config.comparison_periods} períodos",styles["BodyCustom"]),Spacer(1,5),Paragraph("1. Resumen ejecutivo y decisiones",styles["H2Custom"]),Table([[_style_table(Table(kpi,colWidths=[4.7*cm,2.7*cm]),6.0),_findings_box(styles,findings)]],colWidths=[7.6*cm,9.6*cm],style=[("VALIGN",(0,0),(-1,-1),"TOP")]),Spacer(1,4),Image(str(p1),width=17.6*cm,height=8.0*cm),PageBreak(),Paragraph("2. Evolución, alertas y metodología",styles["H2Custom"]),_style_table(Table(trend[:12],colWidths=[5.2*cm,2*cm,2*cm,2.2*cm,2.2*cm,2.3*cm]),5.7),Spacer(1,8)]+_methodology_flowables(styles,risk,rec)+[PageBreak(),Paragraph("3. Análisis técnico de alertas de sensores",styles["H2Custom"]),_sensor_cards(styles,sensor_summary),Spacer(1,7),_sensor_summary_table(styles,sensor_summary),Spacer(1,10),Paragraph("Tractos con eventos de sensores",styles["H2Custom"]),_style_table(Table(equipment_rows,colWidths=[2.6*cm,2.3*cm,1.7*cm,2.3*cm,1.5*cm,2.5*cm,2.2*cm]),5.4),Spacer(1,7),Paragraph("Prioridad alta cuando un tracto registra tres o más eventos o presenta Sensor Tapado y Sensor desalineado en el mismo período.",styles["BodyCustom"]),PageBreak(),Paragraph("4. Conductores prioritarios",styles["H2Custom"]),Image(str(p3),width=17.6*cm,height=6.3*cm),Spacer(1,4)]
    rows=[["Conductor","Eventos","Reinc.","Fatiga NO","Pend.","Cint.","Fuma","Riesgo","Nivel"]]
    for _,r in drivers.iterrows(): rows.append([Paragraph(truncate(r.Conductor,34),styles["SmallCustom"]),int(r.Alertas),int(r.Reincidencias),int(r["Fatiga sin cumplir"]),int(r["Fatiga pendiente"]),int(r["Sin cinturón"]),int(r.Fumando),int(r.Riesgo),r.Nivel])
    elements += [_style_table(Table(rows,colWidths=[5.2*cm,1.1*cm,1.1*cm,1.2*cm,1*cm,1*cm,1*cm,1.1*cm,1.4*cm]),5.0),Spacer(1,5),Paragraph("El orden prioriza fatiga sin cumplimiento, reincidencia, cinturón/fumando y volumen total.",styles["BodyCustom"]),PageBreak(),Paragraph("5. Plan de acción y seguimiento",styles["H2Custom"]),_style_table(Table([[Paragraph(str(c),styles["SmallCustom"]) for c in row] for row in _action_rows(risk,rec,"Transportista")],colWidths=[1.8*cm,4.4*cm,6.4*cm,3.1*cm,1.3*cm]),5.5),Spacer(1,8),Paragraph("Seguimiento del período siguiente",styles["H2Custom"]),Paragraph("Verificar reducción por tipo de alerta, reincidencia de los conductores intervenidos y cierre de fatigas pendientes. Toda acción debe registrar responsable, fecha y resultado.",styles["BodyCustom"]),PageBreak(),Paragraph("6. Detalle completo de eventos",styles["H2Custom"])]
    detail=_detail_dataframe(current); detail_cols=[c for c in ["ID","Fecha","Plataforma","Conductor","Tracto","Patente","Tipo de evento","Cumplimiento fatiga"] if c in detail.columns]; drows=[detail_cols]
    for _,r in detail[detail_cols].iterrows(): drows.append([Paragraph(truncate(r[col],24),styles["TinyCustom"]) for col in detail_cols])
    widths=[1.4,1.7,1.8,4.0,1.5,1.7,3.0,2.0][:len(detail_cols)]; elements += [_style_table(LongTable(drows,colWidths=[x*cm for x in widths],repeatRows=1),4.7)]
    elements = executive_pages(current, previous, config, transportista, risk=risk) + elements + rco_appendix(current, previous)
    doc.build(elements,onFirstPage=_page_branding,onLaterPages=_page_branding)
    return pdf


def generate_reports(data,config):
    data=filter_real_transportistas(data); cutoff=pd.Timestamp(datetime.combine(config.period_end,datetime.max.time())); data=data[data["Fecha"]<=cutoff].copy(); current=_period_slice(data,config.period_start,config.period_end)
    if current.empty: raise ValueError("No existen alertas dentro del período seleccionado.")
    with tempfile.TemporaryDirectory(prefix="copec_reports_") as tmp:
        out=Path(tmp); global_bytes=None; files=[]; global_path=None
        if config.include_global: global_path=generate_global_report(data,config,out); global_bytes=global_path.read_bytes()
        transportistas=sorted(current.Transportista.dropna().unique()) if config.include_transportistas else []
        for index,transportista in enumerate(transportistas,1): files.append(generate_transportista_report(data,transportista,config,out,index))
        buffer=io.BytesIO()
        with zipfile.ZipFile(buffer,"w",zipfile.ZIP_DEFLATED) as archive:
            if global_path: archive.write(global_path,arcname=global_path.name)
            for path in files: archive.write(path,arcname=path.name)
        zip_bytes=buffer.getvalue()
    risk=operational_risk_summary(current); rec=recurrence_summary(current)
    return global_bytes,zip_bytes,{"total_alertas":len(current),"transportistas":len(transportistas),"conductores":driver_count(current),"equipos":current.Tracto.nunique(),"fatiga":int((current.Tipo=="Fatiga").sum()),"riesgo":risk["level"],"reincidencia":rec["rate"],"modo":config.report_mode,"sin_rco":rco_summary(current)["sin_rco"],"pct_sin_rco":rco_summary(current)["pct"]}