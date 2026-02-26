from __future__ import annotations

from io import BytesIO
from typing import Dict, List, Any, Optional

from xml.sax.saxutils import escape
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas as canvas_module


def _p(text: str, style) -> Paragraph:
    safe = escape(text or "").replace("\n", "<br/>")
    return Paragraph(safe, style)


def _meaningful(value: Any) -> bool:
    if value is None:
        return False
    s = str(value).strip()
    if not s:
        return False
    low = s.lower()
    bad = {
        "non pertinente",
        "non applicabile",
        "n/a",
        "na",
        "—",
        "-",
        "nessuna",
        "nessuna / non applicabile",
    }
    if low in bad:
        return False
    if "xxxx" in low:
        return False
    return True


def _kv_table(rows: List[list], col_widths):
    tbl = Table(rows, colWidths=col_widths, hAlign="LEFT")
    tbl.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return tbl


class NumberedCanvas(canvas_module.Canvas):
    """Canvas che consente di stampare 'Pagina X di Y' (richiede 2-pass)."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        super().showPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_page_number(num_pages)
            super().showPage()
        super().save()

    def _draw_page_number(self, page_count: int):
        # footer standard (a destra)
        self.setFont("Helvetica", 9)
        self.drawRightString(200 * mm, 12 * mm, f"Pagina {self._pageNumber} di {page_count}")


def _header_footer(canvas, doc, *, header_left: str = "", footer_left: str = ""):
    canvas.saveState()

    # Header (in alto a sinistra)
    if _meaningful(header_left):
        canvas.setFont("Helvetica", 9)
        canvas.drawString(doc.leftMargin, A4[1] - 12 * mm, header_left)

    # Footer (in basso a sinistra)
    if _meaningful(footer_left):
        canvas.setFont("Helvetica", 9)
        canvas.drawString(doc.leftMargin, 12 * mm, footer_left)

    canvas.restoreState()


def _build_cover_page(story: List[Any], data: Dict[str, Any], styles):
    """Cover ispirata alla relazione tecnica allegata."""

    title_style = ParagraphStyle(
        "cover_title",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=22,
        leading=26,
        alignment=1,  # center
        spaceAfter=10,
    )
    sub_style = ParagraphStyle(
        "cover_sub",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=11,
        leading=14,
        alignment=1,
        spaceAfter=8,
    )
    box_style = ParagraphStyle(
        "cover_box",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=11,
        leading=14,
        alignment=0,
        spaceAfter=6,
    )
    small_center = ParagraphStyle(
        "small_center",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10,
        leading=13,
        alignment=1,
    )

    # Titoli (se non forniti, si ricavano dai campi principali)
    cover_main = data.get("cover_titolo_principale") or "IMPIANTO ELETTRICO"
    cover_up = data.get("cover_titolo_superiore") or "RELAZIONE TECNICA"
    cover_sub = data.get("cover_sottotitolo") or "Relazione tecnico-specialistica"
    oggetto = data.get("oggetto_intervento", "")
    if _meaningful(oggetto):
        cover_sub = oggetto

    # Riga normativa (facoltativa)
    norma = data.get("cover_riga_normativa") or "Ai sensi del D.M. 37/2008 e norme tecniche applicabili"

    story.append(Spacer(1, 18 * mm))
    story.append(_p(cover_up.upper(), small_center))
    story.append(Spacer(1, 6 * mm))
    story.append(_p(cover_main.upper(), title_style))
    story.append(_p(cover_sub, sub_style))
    story.append(_p(norma, sub_style))
    story.append(Spacer(1, 10 * mm))

    # Dati committente e luogo
    comm = data.get("committente_nome", "")
    luogo = data.get("impianto_indirizzo", "")
    if _meaningful(comm) or _meaningful(luogo):
        story.append(_p(f"<b>Committente:</b> {comm}", box_style))
        if _meaningful(luogo):
            story.append(_p(f"{luogo}", box_style))
        story.append(Spacer(1, 8 * mm))

    # Indice capitoli (stile simile al PDF allegato)
    capitoli = [
        "CAPITOLO 1: Premessa",
        "CAPITOLO 2: Riferimenti Legislativi e normativi",
        "CAPITOLO 3: Criterio di progetto degli impianti",
        "CAPITOLO 4: Dati generali e soggetti coinvolti (sintesi)",
        "CAPITOLO 5: Descrizione dell’impianto e opere eseguite",
        "CAPITOLO 6: Dimensionamento (sintesi) / Elenco circuiti",
        "CAPITOLO 7: Verifiche, prove e collaudi",
        "CAPITOLO 8: Esercizio, manutenzione e avvertenze",
        "CAPITOLO 9: Allegati",
    ]
    # Se non c'è criterio esteso o linee/verifiche ecc., non stampiamo voci non presenti
    filtered = []
    if True:
        filtered.append(capitoli[0])
        filtered.append(capitoli[1])
    if _meaningful(data.get("criterio_progetto", "")):
        filtered.append(capitoli[2])
    filtered.append(capitoli[3])
    filtered.append(capitoli[4])
    if data.get("linee"):
        filtered.append(capitoli[5])
    if _meaningful(data.get("verifiche", "")):
        filtered.append(capitoli[6])
    if _meaningful(data.get("manutenzione", "")):
        filtered.append(capitoli[7])
    if _meaningful(data.get("allegati", "")):
        filtered.append(capitoli[8])

    story.append(_p("<b>Sommario</b>", styles["Heading3"]))
    for c in filtered:
        story.append(_p(c, styles["BodyText"]))
    story.append(Spacer(1, 10 * mm))

    # Progettista / tecnico redattore
    progettista = data.get("progettista_blocco", "") or data.get("firma", "")
    if _meaningful(progettista):
        story.append(_p(f"<b>Il progettista:</b><br/>{progettista}", box_style))

    # Riga documento (codice/rev/data) in fondo pagina (facoltativa)
    cod = data.get("cod_progetto", "")
    rev = data.get("rev", "")
    data_doc = data.get("data", "")
    n_doc = data.get("n_doc", "")

    footer_bits = []
    if _meaningful(cod):
        footer_bits.append(f"Cod. Progetto: {cod}")
    elif _meaningful(n_doc):
        footer_bits.append(f"N. documento: {n_doc}")
    if _meaningful(data_doc):
        footer_bits.append(f"Data: {data_doc}")
    if _meaningful(rev):
        footer_bits.append(f"Rev.: {rev}")

    if footer_bits:
        story.append(Spacer(1, 12 * mm))
        story.append(_p(" – ".join(footer_bits), small_center))

    story.append(PageBreak())


def genera_pdf_relazione_bytes(data: Dict[str, Any]) -> bytes:
    """Genera PDF con cover page + contenuti condizionali."""
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

    # Header/Footer: prova a usare il primo rigo del blocco progettista come intestazione
    header_left = ""
    progettista_blocco = data.get("progettista_blocco", "")
    if _meaningful(progettista_blocco):
        header_left = str(progettista_blocco).splitlines()[0].strip()

    # Footer sinistro tipo "Cod. Progetto: ...  Data: ...  Rev.: ..."
    footer_left_parts = []
    if _meaningful(data.get("cod_progetto", "")):
        footer_left_parts.append(f"Cod. Progetto: {data.get('cod_progetto')}")
    elif _meaningful(data.get("n_doc", "")):
        footer_left_parts.append(f"N. documento: {data.get('n_doc')}")
    if _meaningful(data.get("data", "")):
        footer_left_parts.append(f"Data: {data.get('data')}")
    if _meaningful(data.get("rev", "")):
        footer_left_parts.append(f"Rev.: {data.get('rev')}")
    footer_left = "   ".join(footer_left_parts)

    def _first(canvas, d):
        _header_footer(canvas, d, header_left=header_left, footer_left=footer_left)

    def _later(canvas, d):
        _header_footer(canvas, d, header_left=header_left, footer_left=footer_left)

    story: List[Any] = []

    # 0) Cover page
    _build_cover_page(story, data, styles)

    # Disclaimer breve (come nella relazione allegata: chiarire oggetto e limiti)
    disclaimer = data.get("disclaimer_calcoli", "")
    if not _meaningful(disclaimer):
        disclaimer = (
            "Nota: i calcoli riportati (se presenti) sono di sintesi e hanno finalità di supporto alla relazione "
            "e alla Dichiarazione di Conformità; non sostituiscono un progetto esecutivo completo."
        )
    story.append(_p(disclaimer, styles["BodyText"]))
    story.append(Spacer(1, 8))

    # 1) Identificazione documento (rimane, ma ora dopo cover)
    ident = [
        ["DATI IDENTIFICATIVI DOCUMENTO", ""],
        ["Committente", data.get("committente_nome", "")],
        ["Luogo di installazione", data.get("impianto_indirizzo", "")],
        ["Oggetto intervento", data.get("oggetto_intervento", "")],
        ["Tipologia impianto", data.get("tipologia", "")],
        ["Sistema di distribuzione", data.get("sistema", "")],
        ["Tensione/Frequenza", data.get("tensione", "")],
        ["Potenza impegnata / disponibile", data.get("potenza_disp", "")],
        ["Cod. Progetto", data.get("cod_progetto", "")],
        ["N. documento", data.get("n_doc", "")],
        ["Revisione", data.get("rev", "")],
        ["Data", data.get("data", "")],
    ]
    story.append(_kv_table(ident, [55 * mm, 119 * mm]))
    story.append(Spacer(1, 10))

    story.append(_p("Progettista / Tecnico redattore (se applicabile)", styles["Heading2"]))
    story.append(_p(data.get("progettista_blocco", ""), styles["BodyText"]))
    story.append(Spacer(1, 8))

    story.append(_p("CAPITOLO 1 - PREMESSA", styles["Heading2"]))
    story.append(_p(data.get("premessa", ""), styles["BodyText"]))
    story.append(Spacer(1, 8))

    story.append(_p("CAPITOLO 2 - RIFERIMENTI LEGISLATIVI E NORMATIVI", styles["Heading2"]))
    story.append(_p(data.get("norme", ""), styles["BodyText"]))
    story.append(Spacer(1, 8))

    criterio = data.get("criterio_progetto", "")
    if _meaningful(criterio):
        story.append(_p("CAPITOLO 3 - CRITERIO DI PROGETTO DEGLI IMPIANTI", styles["Heading2"]))
        story.append(_p(criterio, styles["BodyText"]))
        story.append(Spacer(1, 10))

    story.append(_p("CAPITOLO 4 - DATI GENERALI E SOGGETTI COINVOLTI (SINTESI)", styles["Heading2"]))
    impresa = data.get("impresa", "")
    if _meaningful(impresa):
        story.append(_p(f"Impresa installatrice: {impresa}", styles["BodyText"]))
    story.append(_p(data.get("dati_tecnici", ""), styles["BodyText"]))
    story.append(Spacer(1, 8))

    story.append(_p("CAPITOLO 5 - DESCRIZIONE DELL’IMPIANTO E OPERE ESEGUITE", styles["Heading2"]))
    story.append(_p(data.get("descrizione_impianto", ""), styles["BodyText"]))
    story.append(Spacer(1, 8))

    conf = data.get("confini", "")
    if _meaningful(conf):
        story.append(_p("5.1.4 Confini dell’intervento e interfacce", styles["Heading3"]))
        story.append(_p(conf, styles["BodyText"]))
        story.append(Spacer(1, 10))

    quadri = data.get("quadri", [])
    if quadri:
        story.append(_p("5.2 Quadri elettrici e distribuzione", styles["Heading3"]))
        story.append(_p("Tabella quadri (compilazione sintetica):", styles["BodyText"]))
        tdata = [
            [_p("Quadro", th), _p("Ubicazione", th), _p("IP", th), _p("Interruttore generale<br/>(tipo/In)", th), _p("Differenziale generale<br/>(tipo/Idn)", th)]
        ]
        for q in quadri:
            tdata.append(
                [
                    _p(str(q.get("Quadro", "")), tc),
                    _p(str(q.get("Ubicazione", "")), tc),
                    _p(str(q.get("IP", "")), tc),
                    _p(str(q.get("Generale", "")), tc),
                    _p(str(q.get("Diff", "")), tc),
                ]
            )
        tbl = Table(tdata, colWidths=[16 * mm, 40 * mm, 12 * mm, 52 * mm, 54 * mm], repeatRows=1, hAlign="LEFT")
        tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]
            )
        )
        story.append(tbl)
        story.append(Spacer(1, 10))

    linee = data.get("linee", [])
    if linee:
        story.append(_p("CAPITOLO 6 - CRITERI DI PROGETTO E DIMENSIONAMENTO (SINTESI)", styles["Heading2"]))
        story.append(_p("6.2 Elenco circuiti, cavi e protezioni", styles["Heading3"]))
        story.append(_p("Di seguito si riporta l’elenco sintetico dei circuiti principali:", styles["BodyText"]))

        tdata = [
            [
                _p("Circuito<br/>/Linea", th),
                _p("Destinazione<br/>/Utilizzo", th),
                _p("Posa<br/>L (m)", th),
                _p("Cavo<br/>(tipo/sezione)", th),
                _p("Protezione<br/>(MT/MTD)", th),
                _p("Differenziale<br/>(tipo/Idn)", th),
                _p("ΔV %", th),
                _p("Esito", th),
            ]
        ]
        for ln in linee:
            posa = (ln.get("Posa", "") or "").strip()
            ll = ln.get("L_m", "")
            posa_len = f"{posa}\n{ll}" if _meaningful(posa) else f"{ll}"
            tdata.append(
                [
                    _p(str(ln.get("Linea", "")), tc),
                    _p(str(ln.get("Uso", "")), tc),
                    _p(str(posa_len), tc),
                    _p(str(ln.get("Cavo", "")), tc),
                    _p(str(ln.get("Protezione", "")), tc),
                    _p(str(ln.get("Diff", "")), tc),
                    _p(str(ln.get("DV_perc", "")), tc),
                    _p(str(ln.get("Esito", "")), tc),
                ]
            )
        colw = [16 * mm, 30 * mm, 24 * mm, 32 * mm, 26 * mm, 26 * mm, 10 * mm, 10 * mm]
        tbl = Table(tdata, colWidths=colw, repeatRows=1, hAlign="LEFT")
        tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]
            )
        )
        story.append(tbl)
        story.append(Spacer(1, 10))

    story.append(_p("5.4 Protezione contro i contatti diretti e indiretti", styles["Heading3"]))
    story.append(_p(data.get("sicurezza", ""), styles["BodyText"]))
    story.append(Spacer(1, 10))

    ver = data.get("verifiche", "")
    if _meaningful(ver):
        story.append(_p("CAPITOLO 7 - VERIFICHE, PROVE E COLLAUDI", styles["Heading2"]))
        story.append(_p(ver, styles["BodyText"]))
        story.append(Spacer(1, 10))

    man = data.get("manutenzione", "")
    if _meaningful(man):
        story.append(_p("CAPITOLO 8 - ESERCIZIO, MANUTENZIONE E AVVERTENZE", styles["Heading2"]))
        story.append(_p(man, styles["BodyText"]))
        story.append(Spacer(1, 10))

    allg = data.get("allegati", "")
    if _meaningful(allg):
        story.append(_p("CAPITOLO 9 - ALLEGATI", styles["Heading2"]))
        story.append(_p(allg, styles["BodyText"]))

    # Firma
    luogo_f = data.get("luogo_firma", "")
    data_f = data.get("data_firma", "")
    firma = data.get("firma", "")
    if _meaningful(luogo_f) or _meaningful(data_f) or _meaningful(firma):
        story.append(Spacer(1, 14))
        if _meaningful(luogo_f) or _meaningful(data_f):
            txt = "Luogo e data: "
            if _meaningful(luogo_f):
                txt += str(luogo_f).strip()
            if _meaningful(data_f):
                txt += f" – {str(data_f).strip()}" if _meaningful(luogo_f) else str(data_f).strip()
            story.append(_p(txt, styles["BodyText"]))
            story.append(Spacer(1, 8))
        if _meaningful(firma):
            story.append(_p(f"Firma e timbro: {firma}", styles["BodyText"]))

    doc.build(
        story,
        onFirstPage=_first,
        onLaterPages=_later,
        canvasmaker=NumberedCanvas,
    )
    return buf.getvalue()
