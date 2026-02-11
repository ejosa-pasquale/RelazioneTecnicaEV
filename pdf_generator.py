from __future__ import annotations

from io import BytesIO
from typing import Dict, List, Any

from xml.sax.saxutils import escape
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

def _p(text: str, style):
    safe = escape(text or "").replace("\n", "<br/>")
    return Paragraph(safe, style)

def _page_number(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 9)
    canvas.drawRightString(200 * mm, 12 * mm, f"Pag. {doc.page}")
    canvas.restoreState()

def _kv_table(rows: List[list], col_widths):
    tbl = Table(rows, colWidths=col_widths, hAlign="LEFT")
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

def _meaningful(value: Any) -> bool:
    if value is None:
        return False
    s = str(value).strip()
    if not s:
        return False
    low = s.lower()
    # parole/placeholder da non stampare
    bad = {
        "non pertinente", "non applicabile", "n/a", "na", "—", "-", "nessuna", "nessuna / non applicabile"
    }
    if low in bad:
        return False
    if "xxxx" in low:
        return False
    return True

def _append_if(story: List[Any], title: str, text: str, styles):
    if _meaningful(text):
        story.append(_p(title, styles["Heading3"]))
        story.append(_p(text, styles["BodyText"]))
        story.append(Spacer(1, 8))

def genera_pdf_relazione_bytes(data: Dict[str, Any]) -> bytes:
    """Genera PDF allineato al template v7 con contenuti condizionali (stampa solo ciò che serve)."""
    buf = BytesIO()
    styles = getSampleStyleSheet()

    th = ParagraphStyle("th", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8, leading=9)
    tc = ParagraphStyle("tc", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=9)

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
        ["Committente", data.get("committente_nome","")],
        ["Luogo di installazione", data.get("impianto_indirizzo","")],
        ["Oggetto intervento", data.get("oggetto_intervento","")],
        ["Tipologia impianto", data.get("tipologia","")],
        ["Sistema di distribuzione", data.get("sistema","")],
        ["Tensione/Frequenza", data.get("tensione","")],
        ["Potenza impegnata / disponibile", data.get("potenza_disp","")],
        ["N. documento", data.get("n_doc","")],
        ["Revisione", data.get("rev","")],
        ["Data", data.get("data","")],
    ]
    story.append(_kv_table(ident, [55*mm, 119*mm]))
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
    impresa = data.get("impresa","")
    if _meaningful(impresa):
        story.append(_p(f"Impresa installatrice: {impresa}", styles["BodyText"]))
    story.append(_p(data.get("dati_tecnici",""), styles["BodyText"]))
    story.append(Spacer(1, 8))

    story.append(_p("CAPITOLO 4 - DESCRIZIONE DELL’IMPIANTO E OPERE ESEGUITE", styles["Heading2"]))
    story.append(_p(data.get("descrizione_impianto",""), styles["BodyText"]))
    story.append(Spacer(1, 8))

    # Confini intervento solo se valorizzati
    conf = data.get("confini","")
    if _meaningful(conf):
        story.append(_p("4.1.4 Confini dell’intervento e interfacce", styles["Heading3"]))
        story.append(_p(conf, styles["BodyText"]))
        story.append(Spacer(1, 10))

    # Quadri solo se presenti
    quadri = data.get("quadri", [])
    if quadri:
        story.append(_p("4.2 Quadri elettrici e distribuzione", styles["Heading3"]))
        story.append(_p("Tabella quadri (compilazione sintetica):", styles["BodyText"]))
        tdata = [
            [_p("Quadro", th), _p("Ubicazione", th), _p("IP", th),
             _p("Interruttore generale<br/>(tipo/In)", th),
             _p("Differenziale generale<br/>(tipo/Idn)", th)]
        ]
        for q in quadri:
            tdata.append([
                _p(str(q.get("Quadro","")), tc),
                _p(str(q.get("Ubicazione","")), tc),
                _p(str(q.get("IP","")), tc),
                _p(str(q.get("Generale","")), tc),
                _p(str(q.get("Diff","")), tc),
            ])
        tbl = Table(tdata, colWidths=[16*mm, 40*mm, 12*mm, 52*mm, 54*mm], repeatRows=1, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),3),
            ("RIGHTPADDING",(0,0),(-1,-1),3),
            ("TOPPADDING",(0,0),(-1,-1),2),
            ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 10))

    # Linee solo se presenti
    linee = data.get("linee", [])
    if linee:
        story.append(_p("CAPITOLO 5 - CRITERI DI PROGETTO E DIMENSIONAMENTO (SINTESI)", styles["Heading2"]))
        story.append(_p("5.2 Elenco circuiti, cavi e protezioni", styles["Heading3"]))
        story.append(_p("Di seguito si riporta l’elenco sintetico dei circuiti principali:", styles["BodyText"]))

        tdata = [[
            _p("Circuito<br/>/Linea", th),
            _p("Destinazione<br/>/Utilizzo", th),
            _p("Posa<br/>L (m)", th),
            _p("Cavo<br/>(tipo/sezione)", th),
            _p("Protezione<br/>(MT/MTD)", th),
            _p("Differenziale<br/>(tipo/Idn)", th),
            _p("ΔV %", th),
            _p("Esito", th),
        ]]
        for ln in linee:
            posa = (ln.get("Posa","") or "").strip()
            ll = ln.get("L_m","")
            posa_len = f"{posa}\n{ll}" if _meaningful(posa) else f"{ll}"
            tdata.append([
                _p(str(ln.get("Linea","")), tc),
                _p(str(ln.get("Uso","")), tc),
                _p(str(posa_len), tc),
                _p(str(ln.get("Cavo","")), tc),
                _p(str(ln.get("Protezione","")), tc),
                _p(str(ln.get("Diff","")), tc),
                _p(str(ln.get("DV_perc","")), tc),
                _p(str(ln.get("Esito","")), tc),
            ])
        colw = [16*mm, 30*mm, 24*mm, 32*mm, 26*mm, 26*mm, 10*mm, 10*mm]
        tbl = Table(tdata, colWidths=colw, repeatRows=1, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),3),
            ("RIGHTPADDING",(0,0),(-1,-1),3),
            ("TOPPADDING",(0,0),(-1,-1),2),
            ("BOTTOMPADDING",(0,0),(-1,-1),2),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 10))

    # Sicurezza: stampa solo se meaningful (di base lo è)
    story.append(_p("5.4 Protezione contro i contatti diretti e indiretti", styles["Heading3"]))
    story.append(_p(data.get("sicurezza",""), styles["BodyText"]))
    story.append(Spacer(1, 10))

    # Verifiche: se tutto è vuoto non stampare (ma di solito serve)
    ver = data.get("verifiche","")
    if _meaningful(ver):
        story.append(_p("CAPITOLO 6 - VERIFICHE, PROVE E COLLAUDI", styles["Heading2"]))
        story.append(_p(ver, styles["BodyText"]))
        story.append(Spacer(1, 10))

    man = data.get("manutenzione","")
    if _meaningful(man):
        story.append(_p("CAPITOLO 7 - ESERCIZIO, MANUTENZIONE E AVVERTENZE", styles["Heading2"]))
        story.append(_p(man, styles["BodyText"]))
        story.append(Spacer(1, 10))

    allg = data.get("allegati","")
    if _meaningful(allg):
        story.append(_p("CAPITOLO 8 - ALLEGATI", styles["Heading2"]))
        story.append(_p(allg, styles["BodyText"]))

    # Firma: stampa solo se valorizzata
    luogo_f = data.get("luogo_firma","")
    data_f = data.get("data_firma","")
    firma = data.get("firma","")
    if _meaningful(luogo_f) or _meaningful(data_f) or _meaningful(firma):
        story.append(Spacer(1, 14))
        if _meaningful(luogo_f) or _meaningful(data_f):
            story.append(_p(f"Luogo e data: {luogo_f} – {data_f}".strip(" –"), styles["BodyText"]))
            story.append(Spacer(1, 8))
        if _meaningful(firma):
            story.append(_p(f"Firma e timbro: {firma}", styles["BodyText"]))

    doc.build(story, onFirstPage=_page_number, onLaterPages=_page_number)
    return buf.getvalue()
