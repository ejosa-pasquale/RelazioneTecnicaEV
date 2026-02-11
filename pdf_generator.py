from __future__ import annotations

from io import BytesIO
from typing import Dict, List, Any

from xml.sax.saxutils import escape
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def _p(text: str, style):
    safe = escape(text).replace("\n", "<br/>")
    return Paragraph(safe, style)

def _page_number(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 9)
    canvas.drawRightString(200 * mm, 12 * mm, f"Pag. {doc.page}")
    canvas.restoreState()

def _kv_table(rows: List[list], col_widths):
    tbl = Table(rows, colWidths=col_widths)
    tbl.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),4),
        ("RIGHTPADDING",(0,0),(-1,-1),4),
        ("TOPPADDING",(0,0),(-1,-1),3),
        ("BOTTOMPADDING",(0,0),(-1,-1),3),
    ]))
    return tbl

def genera_pdf_relazione_bytes(data: Dict[str, Any]) -> bytes:
    """Genera PDF 'Relazione Tecnica - Impianto Elettrico (DiCo)' allineato al template v7."""
    buf = BytesIO()
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title="Relazione Tecnica - Impianto Elettrico (DiCo)",
    )

    story: List[Any] = []

    story.append(_p("RELAZIONE TECNICA - IMPIANTO ELETTRICO (DiCo)", styles["Title"]))
    story.append(_p("Ai sensi del D.M. 37/2008 e norme tecniche applicabili", styles["BodyText"]))
    story.append(Spacer(1, 10))

    ident = [
        ["DATI IDENTIFICATIVI DOCUMENTO", ""],
        ["Committente", data.get("committente_nome","XXXX (Inserire)")],
        ["Luogo di installazione", data.get("impianto_indirizzo","XXXX (Inserire indirizzo completo)")],
        ["Oggetto intervento", data.get("oggetto_intervento","XXXX (Inserire)")],
        ["Tipologia impianto", data.get("tipologia","XXXX (Inserire)")],
        ["Sistema di distribuzione", data.get("sistema","XXXX (Inserire)")],
        ["Tensione/Frequenza", data.get("tensione","XXXX (Inserire)")],
        ["Potenza impegnata / disponibile", data.get("potenza_disp","XXXX (Inserire)")],
        ["N. documento", data.get("n_doc","XXXX (Inserire)")],
        ["Revisione", data.get("rev","00")],
        ["Data", data.get("data","XXXX (Inserire)")],
    ]
    story.append(_kv_table(ident, [55*mm, 120*mm]))
    story.append(Spacer(1, 10))

    story.append(_p("3.3 Progettista / Tecnico redattore (se applicabile)", styles["Heading2"]))
    story.append(_p(data.get("progettista_blocco",""), styles["BodyText"]))
    story.append(Spacer(1, 8))

    story.append(_p("CAPITOLO 1 - PREMESSA", styles["Heading2"]))
    story.append(_p(data.get("premessa",""), styles["BodyText"]))
    story.append(Spacer(1, 8))

    story.append(_p("CAPITOLO 2 - RIFERIMENTI LEGISLATIVI E NORMATIVI", styles["Heading2"]))
    story.append(_p(data.get("norme",""), styles["BodyText"]))
    story.append(Spacer(1, 8))

    story.append(_p("CAPITOLO 3 - DATI GENERALI E SOGGETTI COINVOLTI (SINTESI)", styles["Heading2"]))
    story.append(_p(f"Impresa installatrice: {data.get('impresa','XXXX (Inserire)')}", styles["BodyText"]))
    story.append(_p(data.get("dati_tecnici",""), styles["BodyText"]))
    story.append(Spacer(1, 8))

    story.append(_p("CAPITOLO 4 - DESCRIZIONE DELL’IMPIANTO E OPERE ESEGUITE", styles["Heading2"]))
    story.append(_p(data.get("descrizione_impianto",""), styles["BodyText"]))
    story.append(Spacer(1, 8))

    story.append(_p("4.1.4 Confini dell’intervento e interfacce", styles["Heading3"]))
    story.append(_p(data.get("confini",""), styles["BodyText"]))
    story.append(Spacer(1, 10))

    story.append(_p("4.2 Quadri elettrici e distribuzione", styles["Heading3"]))
    story.append(_p("Tabella quadri (compilazione sintetica):", styles["BodyText"]))
    quadri = data.get("quadri", [])
    if quadri:
        tdata = [["Quadro", "Ubicazione", "IP", "Interruttore generale (tipo/In)", "Differenziale generale (tipo/Idn, se presente)"]]
        for q in quadri:
            tdata.append([q.get("Quadro",""), q.get("Ubicazione",""), q.get("IP",""), q.get("Generale",""), q.get("Diff","")])
        tbl = Table(tdata, colWidths=[18*mm, 45*mm, 14*mm, 55*mm, 45*mm])
        tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("FONTSIZE",(0,0),(-1,-1),8),
        ]))
        story.append(tbl)
    else:
        story.append(_p("—", styles["BodyText"]))
    story.append(Spacer(1, 10))

    story.append(_p("CAPITOLO 5 - CRITERI DI PROGETTO E DIMENSIONAMENTO (SINTESI)", styles["Heading2"]))
    story.append(_p("5.2 Elenco circuiti, cavi e protezioni", styles["Heading3"]))
    story.append(_p("Di seguito si riporta l’elenco sintetico dei circuiti principali:", styles["BodyText"]))
    linee = data.get("linee", [])
    if linee:
        tdata = [["Circuito/Linea", "Destinazione/Utilizzo", "Posa / Lunghezza", "Cavo (tipo / sezione)", "Protezione (MT/MTD)", "Differenziale (tipo/Idn)", "ΔV %", "Esito"]]
        for ln in linee:
            posa = ln.get("Posa","")
            ll = ln.get("L_m","")
            posa_len = f"{posa} - {ll} m" if posa else f"{ll} m"
            tdata.append([
                ln.get("Linea",""),
                ln.get("Uso",""),
                posa_len,
                ln.get("Cavo",""),
                ln.get("Protezione",""),
                ln.get("Diff",""),
                ln.get("DV_perc",""),
                ln.get("Esito",""),
            ])
        tbl = Table(tdata, colWidths=[18*mm, 34*mm, 32*mm, 34*mm, 32*mm, 32*mm, 14*mm, 14*mm])
        tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("FONTSIZE",(0,0),(-1,-1),7.6),
        ]))
        story.append(tbl)
    else:
        story.append(_p("—", styles["BodyText"]))
    story.append(Spacer(1, 10))

    story.append(_p("5.4 Protezione contro i contatti diretti e indiretti", styles["Heading3"]))
    story.append(_p(data.get("sicurezza",""), styles["BodyText"]))
    story.append(Spacer(1, 10))

    story.append(_p("CAPITOLO 6 - VERIFICHE, PROVE E COLLAUDI", styles["Heading2"]))
    story.append(_p(data.get("verifiche",""), styles["BodyText"]))
    story.append(Spacer(1, 10))

    story.append(_p("CAPITOLO 7 - ESERCIZIO, MANUTENZIONE E AVVERTENZE", styles["Heading2"]))
    story.append(_p(data.get("manutenzione",""), styles["BodyText"]))
    story.append(Spacer(1, 10))

    story.append(_p("CAPITOLO 8 - ALLEGATI", styles["Heading2"]))
    story.append(_p(data.get("allegati",""), styles["BodyText"]))

    story.append(Spacer(1, 14))
    story.append(_p("Luogo e data: XXXX (Inserire)", styles["BodyText"]))
    story.append(Spacer(1, 10))
    story.append(_p("Firma e timbro: XXXX (Inserire)", styles["BodyText"]))

    doc.build(story, onFirstPage=_page_number, onLaterPages=_page_number)
    return buf.getvalue()
