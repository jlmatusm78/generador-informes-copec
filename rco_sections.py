"""Identification metrics and readable PDF sections; no changes to source data."""
from pathlib import Path
import re
import unicodedata
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, Table, TableStyle, Spacer, PageBreak, LongTable

NAVY = colors.HexColor('#173A54')
RED = colors.HexColor('#D90016')
LIGHT = colors.HexColor('#F1F5F8')

def normalized(value):
    if pd.isna(value):
        return ''
    return re.sub(r'[^A-Z0-9]', '', ''.join(c for c in unicodedata.normalize('NFD', str(value).upper()) if unicodedata.category(c) != 'Mn'))

def rco_mask(df):
    return df['Conductor'].map(normalized).eq('SINRCO')

def identified_mask(df):
    # Missing names are not silently reclassified as confirmed Sin RCO.
    missing = {'', 'SINRCO', 'SINCONDUCTOR', 'SINIDENTIFICAR', 'NOIDENTIFICADO',
               'SINIDENTIFICACION', 'DESCONOCIDO', 'SINDATO', 'SINDATOS', 'SI', 'NA', 'NAN', 'NONE'}
    return ~df['Conductor'].map(normalized).isin(missing)

def identified_data(df):
    return df.loc[identified_mask(df)].copy()

def driver_keys(df):
    return ['Transportista', 'Conductor'] if 'Transportista' in df.columns else ['Conductor']

def driver_count(df):
    known = identified_data(df)
    return len(known[driver_keys(known)].drop_duplicates())

def rco_summary(df):
    total = len(df)
    missing = int(rco_mask(df).sum())
    known = int(identified_mask(df).sum())
    return dict(total=total, sin_rco=missing, identified=known, other=total-missing-known,
                pct=100*missing/total if total else None,
                identified_pct=100*known/total if total else None)

def fmt_pct(value):
    return 'No aplica' if value is None or pd.isna(value) else f'{value:.1f}%'.replace('.', ',')

def body(text, size=9.4, color=NAVY):
    return Paragraph(text, ParagraphStyle('RCOBody', fontName='Helvetica', fontSize=size,
                     leading=size*1.4, textColor=color, spaceAfter=6))

def title(text):
    return Paragraph(text, ParagraphStyle('RCOTitle', fontName='Helvetica-Bold', fontSize=17,
                     leading=21, textColor=NAVY, spaceAfter=10))

def heading(text):
    return Paragraph(text, ParagraphStyle('RCOHeading', fontName='Helvetica-Bold', fontSize=12,
                     leading=16, textColor=NAVY, spaceBefore=10, spaceAfter=8))

def tabular(headers, rows, widths):
    data = [[body(escape(str(x)), 8.3, colors.white) for x in headers]]
    data += [[body(escape(str(x)), 8.2) for x in row] for row in rows]
    table = LongTable(data, colWidths=widths, repeatRows=1, hAlign='LEFT')
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), NAVY),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT]),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING',(0,0),(-1,-1),7), ('RIGHTPADDING',(0,0),(-1,-1),7),
        ('TOPPADDING',(0,0),(-1,-1),6), ('BOTTOMPADDING',(0,0),(-1,-1),6),
    ]))
    return table

def cards(items):
    cells = []
    for value, label, note in items:
        cells.append([body(f'<b>{escape(str(value))}</b>', 23, RED if label == 'Alertas Sin RCO' else NAVY), body(f'<b>{escape(label)}</b>',9), body(escape(note),8)])
    width = 18*cm/len(cells)
    t = Table([cells],colWidths=[width]*len(cells))
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),LIGHT),('VALIGN',(0,0),(-1,-1),'TOP'),
                         ('BOX',(0,0),(-1,-1),0,colors.white),
                         ('LINEAFTER',(0,0),(-2,-1),7,colors.white),
                         ('LEFTPADDING',(0,0),(-1,-1),10),('RIGHTPADDING',(0,0),(-1,-1),10),
                         ('TOPPADDING',(0,0),(-1,-1),10),('BOTTOMPADDING',(0,0),(-1,-1),10)]))
    return t

def previous_period(data, config):
    duration = (config.week_end-config.week_start).days+1
    end = pd.Timestamp(config.week_start)
    return data[(data['Fecha'] >= end-pd.Timedelta(days=duration)) & (data['Fecha'] < end)].copy()

def executive_pages(current, previous, config, transportista=None, risk=None):
    s, p = rco_summary(current), rco_summary(previous)
    fatigue = int(current['Tipo'].eq('Fatiga').sum())
    fatigue_missing = int((current['Tipo'].eq('Fatiga') & rco_mask(current)).sum())
    change = (f"{(s['total']/p['total']-1)*100:+.1f}%".replace('.', ',') if p['total'] else 'Sin base')
    delta = (f"{s['pct']-p['pct']:+.1f}".replace('.', ',')+' puntos porcentuales'
             if s['pct'] is not None and p['pct'] is not None else 'Sin base de comparación')
    label = f'Informe {config.report_mode.lower()} de alertas' if transportista else f'Informe {config.report_mode.lower()} global COPEC'
    subtitle = escape(transportista) if transportista else 'Consolidado de transportistas incluidos'
    elements = [title(label), body(f'{subtitle}<br/>Período: {config.week_start:%d-%m-%Y} al {config.week_end:%d-%m-%Y}'),
        heading('Resumen ejecutivo'), cards([
            (s['total'],'Alertas totales',f"Período anterior: {p['total']}"),
            (change,'Variación de alertas','Respecto del período anterior'),
            (fatigue,'Alertas de fatiga',f'Incluye {fatigue_missing} Sin RCO')]),
        heading('Identificación de conductores - Sin RCO'),
        body('<b>Sin RCO:</b> el conductor no se identificó en el sistema al momento de la alerta.'),
        cards([(s['sin_rco'],'Alertas Sin RCO',f"De {s['total']} alertas totales"),
               (fmt_pct(s['pct']),'Alertas Sin RCO',f"Anterior: {fmt_pct(p['pct'])}"),
               (fmt_pct(s['identified_pct']),'Con identificación',f"{s['identified']} alertas atribuibles a personas")]),
        Spacer(1,8), body(f'<b>Cambio del % Sin RCO:</b> {delta}.'),
        heading('Lectura práctica'),
        body(f"Se registraron <b>{s['total']} alertas</b>; <b>{s['sin_rco']} están Sin RCO ({fmt_pct(s['pct'])})</b>. "
             f"En el período anterior hubo {p['total']} alertas, de las cuales {p['sin_rco']} estaban Sin RCO. "
             f"El análisis por conductor utiliza <b>{s['identified']} alertas identificadas</b>."),
        body(f"<b>Acción:</b> regularizar la identificación de los {s['sin_rco']} casos Sin RCO, "
             f"priorizando las {fatigue_missing} alertas de fatiga, y reforzar la identificación al iniciar la operación."
             if s['sin_rco'] else '<b>Acción:</b> mantener la identificación al iniciar la operación. No hay alertas marcadas Sin RCO en este período.')]
    if s['other']:
        elements.append(body(f"<b>Otros registros sin identificación:</b> {s['other']}. Se excluyen del análisis por persona y se muestran por separado; no se asumen como Sin RCO.",8.5))
    elements += [PageBreak(),title('Sin RCO: impacto en las mediciones'),
        tabular(['Indicador','Tratamiento de los Sin RCO'],[
            ('Total y tipos de alertas','Se incluyen. La falta de identificación no elimina el evento.'),
            ('Transportista y tracto','Se incluyen cuando la empresa y el equipo están identificados.'),
            ('Conductores y ranking','Se excluyen del análisis por persona. Sin RCO no representa un conductor único.'),
            ('Reincidencia','Se calcula con conductores identificados. Puede subestimar la reincidencia real.'),
            ('Descanso por fatiga','Se evalúa con el registro de detención. Sin RCO no implica automáticamente incumplimiento.')],[5.0*cm,13.0*cm]),
        Spacer(1,10),body('<b>Base de cálculo:</b> % Sin RCO = alertas Sin RCO / total de alertas × 100. '
            'El porcentaje mide alertas, no personas ni viajes. Sin alertas se muestra “No aplica”.'),
        body('<b>Reincidencia:</b> conductores identificados con 2 o más alertas / total de conductores identificados con alertas. '
             'En el consolidado se distingue a cada conductor por transportista. No se suman los Sin RCO como una persona.'),
        body('<b>Comparación:</b> se utiliza la semana completa anterior o el mes calendario anterior, según el modo del informe. '
             'Si no registra alertas, el porcentaje de variación se muestra sin base de comparación.'),
        body('<b>Fatiga:</b> el cumplimiento corresponde a detenciones confirmadas “SI” / total de alertas de fatiga. '
             'Los registros “NO” y pendientes se muestran separados. Si no hay fatiga, no aplica.'),PageBreak()]
    if risk is not None:
        elements.insert(-1, body('<b>Riesgo operacional:</b> las alertas Sin RCO conservan sus puntos según el tipo de evento y el registro de cumplimiento. La falta de identificación no agrega puntos por sí sola. Se excluyen de la regla de reincidencia de fatiga sin cumplimiento por persona. El total de puntos y los umbrales de riesgo se mantienen.'))
    return elements

def rco_appendix(current, previous, global_report=False):
    cases = current.loc[rco_mask(current)].copy()
    parts = [PageBreak(), title('Sin RCO: seguimiento de casos')]
    if global_report:
        rows = []
        for name, group in current.groupby('Transportista',sort=True):
            s = rco_summary(group); p = rco_summary(previous[previous['Transportista'].eq(name)])
            delta = f"{s['pct']-p['pct']:+.1f} pp".replace('.', ',') if p['pct'] is not None else 'Sin base'
            rows.append((name,s['total'],s['sin_rco'],fmt_pct(s['pct']),delta))
        parts += [heading('Desglose por transportista'),tabular(['Transportista','Alertas','Sin RCO','% Sin RCO','Cambio'],rows,[8*cm,2*cm,2*cm,2*cm,4*cm]),Spacer(1,10)]
    parts += [heading(f'Casos por regularizar: {len(cases)}')]
    if cases.empty:
        return parts+[body('No hay alertas marcadas Sin RCO en el período.')]
    parts += [body('Detalle completo. Priorizar fatiga; identificar al conductor y validar los datos antes de actualizar el análisis por persona.',9)]
    cases['_priority'] = ~cases['Tipo'].eq('Fatiga')
    cases = cases.sort_values(['_priority','Fecha'])
    rows = []
    for _, r in cases.iterrows():
        event = r.get('Fecha/Hora de evento')
        when = str(event) if pd.notna(event) and str(event).strip() else r['Fecha'].strftime('%d-%m-%Y')
        action = 'Identificar y verificar descanso.' if r['Tipo']=='Fatiga' else 'Identificar y realizar retroalimentación.'
        row = [when,r['Tracto'],r['Tipo'],action]
        if global_report:
            row.insert(1,r['Transportista'])
        rows.append(row)
    widths = [3.0*cm,4.3*cm,2.0*cm,3.0*cm,5.7*cm] if global_report else [3.4*cm,2.6*cm,4.0*cm,8.0*cm]
    headers = ['Fecha / hora disponible','Tracto','Alerta','Acción']
    if global_report:
        headers.insert(1,'Transportista')
    parts.append(tabular(headers,rows,widths))
    return parts

def branded_page(canvas, doc):
    canvas.saveState()
    logo = Path(__file__).parent/'assets'/'copec.png'
    # Render original image proportionally; clipping removes only its white margins.
    scale = 110/679
    x,y = 1.1*cm, 28.2*cm
    clip = canvas.beginPath(); clip.rect(x,y,110,26); canvas.clipPath(clip,stroke=0)
    canvas.drawImage(str(logo),x-94*scale,y-(650-397)*scale,width=866*scale,height=650*scale)
    canvas.restoreState(); canvas.saveState()
    canvas.setFillColor(NAVY); canvas.setFont('Helvetica',9)
    canvas.drawRightString(19.9*cm,28.6*cm,'TORRE DE CONTROL')
    canvas.setStrokeColor(RED); canvas.setLineWidth(1.3)
    canvas.line(1.1*cm,27.95*cm,19.9*cm,27.95*cm)
    canvas.setFont('Helvetica',7); canvas.setFillColor(colors.HexColor('#526372'))
    canvas.drawString(1.1*cm,.8*cm,'Torre de Control COPEC - Informe automático')
    canvas.drawRightString(19.9*cm,.8*cm,f'Página {doc.page}')
    canvas.restoreState()
