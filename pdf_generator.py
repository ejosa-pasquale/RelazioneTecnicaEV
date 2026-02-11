from __future__ import annotations

from io import BytesIO
from typing import Dict, List, Any

from xml.sax.saxutils import escape
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
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

def genera_pdf_relazione_bytes(data: Dict[str, Any]) -> bytes:
    """
    Genera PDF 'Relazione Tecnica per DiCo – Impianto Elettrico' con contenuto descrittivo + tabelle sintetiche.
    """
    buf = BytesIO()
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title="Relazione Tecnica – DiCo – Impianto Elettrico",
    )

    story: List[Any] = []
    story.append(_p("RELAZIONE TECNICA – IMPIANTO ELETTRICO (Allegato alla DiCo)", styles["Title"]))
    story.append(Spacer(1, 6))
    story.append(_p(f"Committente: {data['committente_nome']}", styles["BodyText"]))
    story.append(_p(f"Ubicazione: {data['impianto_indirizzo']}", styles["BodyText"]))
    story.append(_p(f"Data: {data.get('data','XXXX (Inserire)')}", styles["BodyText"]))
    story.append(Spacer(1, 10))

    # Progettista
    story.append(_p("PROGETTISTA / REDATTORE", styles["Heading2"]))
    story.append(_p(data["progettista_blocco"], styles["BodyText"]))
    story.append(Spacer(1, 8))

    # Premessa e scopo
    story.append(_p("1. PREMESSA E SCOPO", styles["Heading2"]))
    story.append(_p(data["premessa"], styles["BodyText"]))
    story.append(Spacer(1, 6))

    story.append(_p("2. RIFERIMENTI NORMATIVI", styles["Heading2"]))
    story.append(_p(data["norme"], styles["BodyText"]))
    story.append(Spacer(1, 6))

    story.append(_p("3. DATI TECNICI ESSENZIALI", styles["Heading2"]))
    story.append(_p(data["dati_tecnici"], styles["BodyText"]))
    story.append(Spacer(1, 6))

    story.append(_p("4. DESCRIZIONE DELL’IMPIANTO", styles["Heading2"]))
    story.append(_p(data["descrizione_impianto"], styles["BodyText"]))
    story.append(Spacer(1, 8))

    # Confini intervento
    story.append(_p("4.1 CONFINE INTERVENTO E INTERFACCE", styles["Heading3"]))
    story.append(_p(data["confini"], styles["BodyText"]))
    story.append(Spacer(1, 8))

    # Tabelle quadri
    story.append(_p("4.2 QUADRI E DISTRIBUZIONE", styles["Heading3"]))
    story.append(_p("Tabella quadri (sintesi):", styles["BodyText"]))
    quadri = data.get("quadri", [])
    if quadri:
        tdata = [["Quadro", "Ubicazione", "IP", "Generale", "Differenziale generale"]]
        for q in quadri:
            tdata.append([q.get("Quadro",""), q.get("Ubicazione",""), q.get("IP",""), q.get("Generale",""), q.get("Diff","")])
        tbl = Table(tdata, colWidths=[22*mm, 45*mm, 15*mm, 45*mm, 55*mm])
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

    # Tabelle linee
    story.append(_p("4.3 LINEE E CIRCUITI", styles["Heading3"]))
    story.append(_p("Tabella linee/circuiti (cavo e protezioni):", styles["BodyText"]))
    linee = data.get("linee", [])
    if linee:
        tdata = [["Linea", "Uso", "L (m)", "Cavo", "Protezione", "Diff (tipo/Idn)", "ΔV %", "Esito"]]
        for ln in linee:
            tdata.append([
                ln.get("Linea",""),
                ln.get("Uso",""),
                str(ln.get("L_m","")),
                ln.get("Cavo",""),
                ln.get("Protezione",""),
                ln.get("Diff",""),
                ln.get("DV_perc",""),
                ln.get("Esito",""),
            ])
        tbl = Table(tdata, colWidths=[18*mm, 28*mm, 12*mm, 34*mm, 32*mm, 32*mm, 14*mm, 16*mm])
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

    # Protezioni / Terra / SPD
    story.append(_p("5. SICUREZZA E PROTEZIONI (SINTESI)", styles["Heading2"]))
    story.append(_p(data["sicurezza"], styles["BodyText"]))
    story.append(Spacer(1, 8))

    story.append(_p("6. VERIFICHE E PROVE", styles["Heading2"]))
    story.append(_p(data["verifiche"], styles["BodyText"]))
    story.append(Spacer(1, 10))

    story.append(_p("7. MANUTENZIONE E RACCOMANDAZIONI", styles["Heading2"]))
    story.append(_p(data["manutenzione"], styles["BodyText"]))
    story.append(Spacer(1, 10))

    story.append(_p("8. ALLEGATI", styles["Heading2"]))
    allegati = data.get("allegati", "")
    story.append(_p(allegati, styles["BodyText"]))

    # Firma
    story.append(Spacer(1, 14))
    story.append(_p("Firma e Timbro (se previsto)", styles["BodyText"]))
    story.append(Spacer(1, 18))
    story.append(_p(data.get("firma","Ing. Pasquale Senese"), styles["BodyText"]))

    doc.build(story, onFirstPage=_page_number, onLaterPages=_page_number)
    return buf.getvalue()
