from __future__ import annotations

"""PDF generator.

Obiettivi della versione:
- Cover page in stile "relazione tecnico-specialistica" (3 riquadri: titolo/indice/firma)
- Pagina "Elenco delle revisioni" (subito dopo la cover)
- Header/Footer e numerazione "Pagina X di Y" su tutte le pagine successive
- Wording e struttura più legali/rigorosi (capitoli allineati al template: 1..6)
- Contenuti condizionali: stampa solo sezioni significative

Nota: la cover riprende l'impostazione a tre riquadri del PDF campione.
"""

from io import BytesIO
from typing import Dict, List, Any, Optional

from xml.sax.saxutils import escape

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    Flowable,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


def _p(text: str, style):
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
    if "xxxx" in low or "inserire" in low or "da compil" in low or "da defin" in low:
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


def _first_nonempty_line(text: str) -> str:
    for ln in (text or "").splitlines():
        ln = ln.strip()
        if ln:
            return ln
    return ""


class _NumberedCanvas(canvas.Canvas):
    """Canvas con numerazione 'Pagina X di Y' senza duplicare le pagine."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states: List[dict] = []

    def showPage(self):
        # Salva lo stato della pagina corrente e prepara la successiva
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_page_number(page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def _draw_page_number(self, page_count: int):
        self.saveState()
        self.setFont("Helvetica", 9)
        self.setFillColor(colors.grey)
        self.drawRightString(200 * mm, 10 * mm, f"Pagina {self._pageNumber} di {page_count}")
        self.setFillColor(colors.black)
        self.restoreState()



def _build_indice_items(_: Dict[str, Any]) -> List[str]:
    # Struttura capitoli "editoriale". L'indice è volutamente sintetico:
    # i sotto-paragrafi vengono resi nel testo solo se significativi.
    return [
        "CAPITOLO 1: Premessa",
        "CAPITOLO 2: Riferimenti legislativi e normativi",
        "CAPITOLO 3: Criteri di progetto degli impianti",
        "CAPITOLO 4: Soluzione progettuale adottata (dati tecnici, EVSE, quadri e linee)",
        "CAPITOLO 5: Sicurezza, verifiche e manutenzione",
        "CAPITOLO 6: Allegati e documentazione",
    ]


def _evse_table(data: Dict[str, Any], styles) -> Optional[Table]:
    evse = data.get("evse") or []
    # drop righe vuote (Marca/Modello entrambi non significativi)
    filtered = []
    for r in evse:
        marca = str(r.get("Marca", "")).strip()
        modello = str(r.get("Modello", "")).strip()
        if not _meaningful(marca) and not _meaningful(modello):
            continue
        filtered.append(r)
    if not filtered:
        return None

    th = ParagraphStyle("evth", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8, leading=9)
    tc = ParagraphStyle("evtc", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=9)

    tdata = [[
        _p("Tipo", th),
        _p("Marca/Modello", th),
        _p("Punti", th),
        _p("Potenza", th),
        _p("Alim.", th),
        _p("Connettore/Modo", th),
        _p("IP/IK", th),
        _p("RCD", th),
        _p("Note", th),
    ]]
    for r in filtered:
        marca = str(r.get('Marca', '')).strip()
        modello = str(r.get('Modello', '')).strip()
        marca_modello = (f"{marca} {modello}").strip()
        tdata.append([
            _p(str(r.get("Tipo", "")), tc),
            _p(marca_modello, tc),
            _p(str(r.get("N. punti", "")), tc),
            _p(f"{r.get('Potenza (kW)','')} kW", tc),
            _p(str(r.get("Alimentazione", "")), tc),
            _p(f"{r.get('Connettore','')} / {r.get('Modo ricarica','')}", tc),
            _p(str(r.get("IP/IK", "")), tc),
            _p(str(r.get("RCD richiesto", "")), tc),
            _p(str(r.get("Note", "")), tc),
        ])

    tbl = Table(
        tdata,
        colWidths=[14 * mm, 35 * mm, 10 * mm, 16 * mm, 16 * mm, 26 * mm, 16 * mm, 28 * mm, 23 * mm],
        repeatRows=1,
        hAlign="LEFT",
    )
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return tbl


def _photo_flowables(data: Dict[str, Any]) -> List[Flowable]:
    photos = data.get("foto") or []
    # massimo 6 per non gonfiare il PDF
    photos = [p for p in photos if isinstance(p, dict) and p.get("bytes")][:6]
    if not photos:
        return []

    flows: List[Flowable] = []
    flows.append(Spacer(1, 4))
    # tabella 2 colonne
    cells = []
    for p in photos:
        try:
            img = ImageReader(BytesIO(p["bytes"]))
            iw, ih = img.getSize()
            # fit in box
            box_w = 80 * mm
            box_h = 55 * mm
            scale = min(box_w / iw, box_h / ih)
            w = iw * scale
            h = ih * scale
            # canvas Image can't be directly inserted; use reportlab.platypus Image
            from reportlab.platypus import Image as RLImage

            im = RLImage(BytesIO(p["bytes"]), width=w, height=h)
            caption = Paragraph(escape(p.get("name", "")), getSampleStyleSheet()["BodyText"])
            cells.append([im, caption])
        except Exception:
            continue

    # build rows
    row = []
    rows = []
    for item in cells:
        row.append(item)
        if len(row) == 2:
            rows.append(row)
            row = []
    if row:
        # pad
        row.append([Paragraph("", getSampleStyleSheet()["BodyText"])])
        rows.append(row)

    # convert nested lists into Table with inner tables per cell
    cell_flow_rows = []
    for r in rows:
        out_row = []
        for cell in r:
            if isinstance(cell, list) and len(cell) == 2:
                inner = Table([[cell[0]], [cell[1]]], colWidths=[80 * mm])
                inner.setStyle(TableStyle([
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ]))
                out_row.append(inner)
            else:
                out_row.append(Paragraph("", getSampleStyleSheet()["BodyText"]))
        cell_flow_rows.append(out_row)

    tbl = Table(cell_flow_rows, colWidths=[86 * mm, 86 * mm], hAlign="LEFT")
    tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    flows.append(tbl)
    flows.append(Spacer(1, 6))
    return flows



class EngineeringCoverPage(Flowable):
    """Cover page tipica per documenti di ingegneria con title-block e spazio timbro.

    Layout:
    - Titolo documento al centro-alto
    - Nome progetto in evidenza
    - Title-block tecnico in basso a destra (Cod. progetto / N. doc / Rev / Data / Progettista / Committente)
    - Riquadro timbro/firma in basso a sinistra (con immagine opzionale PNG)
    """

    def __init__(self, data: Dict[str, Any]):
        super().__init__()
        self.data = data

    def wrap(self, availWidth, availHeight):
        return availWidth, availHeight

    def _first_line(self, s: Any) -> str:
        if not s:
            return ""
        if isinstance(s, str):
            return s.strip().split("\n")[0].strip()
        return str(s).strip()

    def draw(self):
        c = self.canv
        # Nota: questo Flowable viene disegnato all'interno del frame (origine = margini).
        # Per ottenere una cover "a pagina intera" su A4 verticale, trasliamo
        # l'origine al vero (0,0) della pagina.
        width, height = A4

        # Deve essere coerente con i margini del SimpleDocTemplate (vedi genera_pdf_relazione_bytes)
        left_margin = 18 * mm
        bottom_margin = 18 * mm
        c.saveState()
        c.translate(-left_margin, -bottom_margin)

        margin_x = 18 * mm
        margin_top = 18 * mm
        margin_bottom = 18 * mm

        # ---- dati
        titolo = self.data.get("titolo_cover") or "RELAZIONE TECNICO-SPECIALISTICA"
        nome_progetto = self.data.get("nome_progetto") or self.data.get("oggetto_intervento") or "Progetto: ________"
        sottotitolo = self.data.get("sottotitolo_cover") or "IMPIANTO ELETTRICO"
        committente = self.data.get("committente_nome") or ""
        indirizzo = self.data.get("impianto_indirizzo") or self.data.get("luogo_intervento") or ""
        cod_progetto = self.data.get("cod_progetto") or ""
        n_doc = self.data.get("n_documento") or self.data.get("n_doc") or ""
        rev = self.data.get("revisione") or self.data.get("rev") or ""
        data_doc = self.data.get("data_documento") or self.data.get("data_doc") or ""
        progettista = (self.data.get("progettista_nome") or self._first_line(self.data.get("progettista_blocco")) 
                       or self._first_line(self.data.get("firma")))

        # ---- stile base
        c.setStrokeColor(colors.black)
        c.setFillColor(colors.black)

        # Titoli (centro pagina, stile "engineering")
        c.setFont("Times-Bold", 22)
        c.drawCentredString(width / 2, height - margin_top - 18 * mm, titolo)

        c.setFont("Times-Bold", 18)
        c.drawCentredString(width / 2, height - margin_top - 32 * mm, nome_progetto)

        c.setFont("Times-Roman", 12)
        c.drawCentredString(width / 2, height - margin_top - 42 * mm, sottotitolo)

        # Riga info (committente / indirizzo)
        info_y = height - margin_top - 58 * mm
        c.setFont("Times-Bold", 10)
        c.drawString(margin_x, info_y, "Committente:")
        c.setFont("Times-Roman", 10)
        c.drawString(margin_x + 26*mm, info_y, committente[:95])

        c.setFont("Times-Bold", 10)
        c.drawString(margin_x, info_y - 6*mm, "Luogo:")
        c.setFont("Times-Roman", 10)
        c.drawString(margin_x + 26*mm, info_y - 6*mm, indirizzo[:95])

        # Linea di separazione
        c.setLineWidth(1)
        c.line(margin_x, info_y - 14*mm, width - margin_x, info_y - 14*mm)

        # ---- title-block (basso)
        block_h = 52 * mm
        block_w = 92 * mm
        block_x = width - margin_x - block_w
        block_y = margin_bottom

        # Riquadro timbro (basso sinistra)
        stamp_w = 82 * mm
        stamp_h = 52 * mm
        stamp_x = margin_x
        stamp_y = margin_bottom
        c.setLineWidth(1)
        c.rect(stamp_x, stamp_y, stamp_w, stamp_h)

        c.setFont("Times-Bold", 9)
        c.drawString(stamp_x + 3*mm, stamp_y + stamp_h - 5*mm, "Spazio timbro / firma")

        timbro_bytes = self.data.get("timbro_bytes") or self.data.get("timbro_image_bytes") or None
        if timbro_bytes:
            try:
                img = ImageReader(BytesIO(timbro_bytes))
                # area immagine con margini
                c.drawImage(img, stamp_x + 3*mm, stamp_y + 3*mm, stamp_w - 6*mm, stamp_h - 12*mm, preserveAspectRatio=True, anchor='c')
            except Exception:
                pass

        # Title-block tecnico
        c.setLineWidth(1)
        c.rect(block_x, block_y, block_w, block_h)

        # griglia interna: 6 righe
        rows = 6
        row_h = block_h / rows
        for i in range(1, rows):
            y = block_y + i * row_h
            c.line(block_x, y, block_x + block_w, y)

        # 2 colonne (label / value)
        split = block_x + 28 * mm
        c.line(split, block_y, split, block_y + block_h)

        labels = ["Cod. Progetto", "N. Documento", "Revisione", "Data", "Progettista", "Committente"]
        values = [cod_progetto, n_doc, rev, data_doc, progettista, committente]

        c.setFont("Times-Bold", 8.5)
        c.setFillColor(colors.black)
        for i, (lab, val) in enumerate(zip(labels, values)):
            y_text = block_y + block_h - (i + 0.7) * row_h
            c.drawString(block_x + 2*mm, y_text, lab)
            c.setFont("Times-Roman", 8.5)
            c.drawString(split + 2*mm, y_text, (val or "")[:40])
            c.setFont("Times-Bold", 8.5)

        # Nota legale minima
        note = self.data.get("disclaimer_cover") or "Documento emesso a supporto della DiCo ex D.M. 37/08; eventuali aggiornamenti normativi successivi non sono inclusi."
        c.setFont("Times-Roman", 8)
        c.drawString(margin_x, block_y + block_h + 6*mm, note[:120])

        c.restoreState()


class LegacyCoverPage(Flowable):
    """Cover a riquadri (titolo / indice / firma) - legacy."""

    def __init__(self, data: Dict[str, Any]):
        super().__init__()
        self.data = data

    def wrap(self, availWidth, availHeight):
        return availWidth, availHeight

    def draw(self):
        c = self.canv
        width, height = A4

        # Anche questa cover è un Flowable: riportiamo l'origine al (0,0) pagina
        left_margin = 18 * mm
        bottom_margin = 18 * mm
        c.saveState()
        c.translate(-left_margin, -bottom_margin)

        left = 20 * mm
        right = 20 * mm
        top = height - 22 * mm
        bottom = 22 * mm
        w = width - left - right

        box1_h = 92 * mm
        box2_h = 45 * mm
        box3_h = (top - bottom) - box1_h - box2_h - 12 * mm

        y1_top = top
        y1_bot = y1_top - box1_h
        y2_top = y1_bot - 6 * mm
        y2_bot = y2_top - box2_h
        y3_top = y2_bot - 6 * mm
        y3_bot = bottom

        c.setLineWidth(1)
        c.rect(left, y1_bot, w, box1_h)
        c.rect(left, y2_bot, w, box2_h)
        c.rect(left, y3_bot, w, box3_h)

        titolo_grande = (self.data.get("titolo_cover") or "RELAZIONE TECNICO-SPECIALISTICA").upper()
        sottotitolo = self.data.get("sottotitolo_cover") or self.data.get("oggetto_intervento") or ""
        committente = self.data.get("committente_nome") or ""
        luogo = self.data.get("impianto_indirizzo") or ""

        # Titolo grande (Times-Bold, centrato, spezzato su più righe)
        c.setFont("Times-Bold", 20)
        words = titolo_grande.split()
        lines: List[str] = []
        cur = ""
        for w0 in words:
            test = (cur + " " + w0).strip()
            if len(test) > 42 and cur:
                lines.append(cur)
                cur = w0
            else:
                cur = test
        if cur:
            lines.append(cur)

        y = y1_top - 12 * mm
        for ln in lines[:6]:
            c.drawCentredString(left + w / 2, y, ln)
            y -= 9 * mm

        c.setFont("Times-Roman", 14)
        y -= 2 * mm
        if _meaningful(sottotitolo):
            c.drawCentredString(left + w / 2, y, str(sottotitolo))
            y -= 8 * mm

        c.setFont("Times-Roman", 13)
        c.drawCentredString(left + w / 2, y, "IMPIANTO ELETTRICO")
        y -= 6 * mm
        c.setFont("Times-Roman", 12)
        c.drawCentredString(left + w / 2, y, "RELAZIONE TECNICO-SPECIALISTICA")
        y -= 10 * mm

        c.setFont("Times-Roman", 12)
        if _meaningful(committente):
            c.drawCentredString(left + w / 2, y, f"Committente: {committente}")
            y -= 6 * mm
        if _meaningful(luogo):
            luogo_txt = str(luogo)
            max_len = 70
            luogo_lines = [luogo_txt[i : i + max_len] for i in range(0, len(luogo_txt), max_len)]
            for ll in luogo_lines[:2]:
                c.drawCentredString(left + w / 2, y, ll)
                y -= 6 * mm

        # Indice con checkbox
        c.setFont("Times-Roman", 13)
        idx = _build_indice_items(self.data)
        x0 = left + 10 * mm
        y = y2_top - 10 * mm
        for it in idx:
            c.rect(x0, y - 3 * mm, 3.5 * mm, 3.5 * mm)
            c.drawString(x0 + 7 * mm, y - 2 * mm, it)
            y -= 7.5 * mm

        # Firma
        progettista = (
            self.data.get("progettista_nome")
            or _first_nonempty_line(self.data.get("progettista_blocco", ""))
            or self.data.get("firma")
            or ""
        )
        c.setFont("Times-Roman", 14)
        c.drawCentredString(left + w / 2, y3_top - 18 * mm, "Il progettista:")
        c.drawCentredString(left + w / 2, y3_top - 28 * mm, str(progettista))

        # Timbro/firma: se fornito PNG, lo disegna; altrimenti placeholder
        stamp_w = 70 * mm
        stamp_h = 45 * mm
        sx = left + (w - stamp_w) / 2
        sy = y3_bot + 16 * mm
        png_bytes = self.data.get("timbro_png")
        if png_bytes:
            try:
                img = ImageReader(BytesIO(png_bytes))
                c.drawImage(img, sx, sy, width=stamp_w, height=stamp_h, preserveAspectRatio=True, mask='auto')
            except Exception:
                png_bytes = None
        if not png_bytes:
            c.setLineWidth(0.5)
            c.rect(sx, sy, stamp_w, stamp_h)
            c.setFont("Helvetica-Oblique", 9)
            c.setFillColor(colors.grey)
            c.drawCentredString(left + w / 2, sy + stamp_h / 2, "Spazio firma/timbro")
            c.setFillColor(colors.black)

        # Metadati (riga bassa)
        cod = self.data.get("cod_progetto", "")
        rev = self.data.get("rev", "")
        data_doc = self.data.get("data", "")
        meta = " · ".join(
            [
                x
                for x in [
                    f"Cod. {cod}" if _meaningful(cod) else "",
                    f"Rev. {rev}" if _meaningful(rev) else "",
                    f"Data {data_doc}" if _meaningful(data_doc) else "",
                ]
                if x
            ]
        )
        if meta:
            c.setFont("Helvetica", 9)
            c.setFillColor(colors.grey)
            c.drawCentredString(left + w / 2, y3_bot + 8 * mm, meta)
            c.setFillColor(colors.black)

        c.restoreState()


def _draw_header_footer(c: canvas.Canvas, doc, data: Dict[str, Any]):
    """Header/Footer per pagine successive alla cover."""
    width, height = A4
    left = doc.leftMargin
    right = width - doc.rightMargin

    titolo = data.get("header_titolo") or "Relazione Tecnico-Specialistica"
    cod = data.get("cod_progetto")
    rev = data.get("rev")
    data_doc = data.get("data")

    c.saveState()
    c.setFont("Helvetica", 9)
    c.setFillColor(colors.grey)

    header = titolo
    if _meaningful(cod):
        header = f"{titolo} · Cod. {cod}"
    c.drawString(left, height - 12 * mm, header)

    c.setStrokeColor(colors.lightgrey)
    c.setLineWidth(0.3)
    c.line(left, 15 * mm, right, 15 * mm)

    meta = " · ".join(
        [x for x in [f"Rev. {rev}" if _meaningful(rev) else "", f"Data {data_doc}" if _meaningful(data_doc) else ""] if x]
    )
    if meta:
        c.drawString(left, 10 * mm, meta)

    c.restoreState()


def _revision_table(data: Dict[str, Any], styles) -> Optional[Table]:
    revs = data.get("revisioni") or []
    if not revs:
        # default minimo
        rev = data.get("rev", "00")
        dt = data.get("data", "")
        revs = [{"Rev": str(rev), "Data": str(dt), "Descrizione": "Emissione documento"}]

    th = ParagraphStyle("rth", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, leading=10)
    tc = ParagraphStyle("rtc", parent=styles["Normal"], fontName="Helvetica", fontSize=9, leading=10)

    tdata = [[_p("Rev.", th), _p("Data", th), _p("Descrizione", th)]]
    for r in revs:
        tdata.append([
            _p(str(r.get("Rev", "")), tc),
            _p(str(r.get("Data", "")), tc),
            _p(str(r.get("Descrizione", "")), tc),
        ])

    tbl = Table(tdata, colWidths=[18 * mm, 30 * mm, 126 * mm], hAlign="LEFT", repeatRows=1)
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return tbl


def genera_pdf_relazione_bytes(data: Dict[str, Any]) -> bytes:
    buf = BytesIO()
    styles = getSampleStyleSheet()

    th = ParagraphStyle("th", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8, leading=9)
    tc = ParagraphStyle("tc", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=9)

    h1 = ParagraphStyle("H1", parent=styles["Heading1"], spaceBefore=6, spaceAfter=8)
    h2 = ParagraphStyle("H2", parent=styles["Heading2"], spaceBefore=8, spaceAfter=6)
    h3 = ParagraphStyle("H3", parent=styles["Heading3"], spaceBefore=6, spaceAfter=4)

    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=20 * mm,
        bottomMargin=18 * mm,
        title="Relazione Tecnico-Specialistica",
    )

    story: List[Any] = []

    # 1) COVER
    cover_style = (data.get('cover_style') or 'engineering').lower()
    story.append(EngineeringCoverPage(data) if cover_style.startswith('eng') else LegacyCoverPage(data))
    story.append(PageBreak())

    # 2) REVISIONI
    story.append(_p("ELENCO DELLE REVISIONI", h1))
    rt = _revision_table(data, styles)
    if rt:
        story.append(rt)
        story.append(Spacer(1, 10))

    # 2.1) Dati identificativi documento
    ident = [
        ["DATI IDENTIFICATIVI DOCUMENTO", ""],
        ["Committente", data.get("committente_nome", "")],
        ["Luogo di installazione", data.get("impianto_indirizzo", "")],
        ["Oggetto intervento", data.get("oggetto_intervento", "")],
        ["Tipologia impianto", data.get("tipologia", "")],
        ["Sistema di distribuzione", data.get("sistema", "")],
        ["Tensione/Frequenza", data.get("tensione", "")],
        ["Potenza impegnata / disponibile", data.get("potenza_disp", "")],
        ["Cod. progetto", data.get("cod_progetto", "")],
        ["N. documento", data.get("n_doc", "")],
        ["Revisione", data.get("rev", "")],
        ["Data", data.get("data", "")],
    ]
    story.append(_kv_table(ident, [55 * mm, 119 * mm]))
    story.append(Spacer(1, 10))

    # 2.2) Dati progettista (se forniti)
    progettista_blocco = data.get("progettista_blocco", "")
    if _meaningful(progettista_blocco):
        story.append(_p("TECNICO PROGETTISTA / REDATTORE", h2))
        story.append(_p(progettista_blocco, styles["BodyText"]))
        story.append(Spacer(1, 10))

    # 2.3) Impresa installatrice (se fornita)
    impresa = data.get("impresa", "")
    if _meaningful(impresa):
        story.append(_p("IMPRESA INSTALLATRICE (dati dichiarati)", h2))
        imp_rows = [
            ["Ragione sociale", impresa],
            ["Sede legale", data.get("impresa_sede", "")],
            ["P.IVA / C.F.", data.get("impresa_piva", "")],
            ["CCIAA / REA", data.get("impresa_rea", "")],
            ["Responsabile tecnico", data.get("impresa_resp", "")],
            ["Recapiti", data.get("impresa_cont", "")],
        ]
        story.append(_kv_table([["Dato", "Valore"]] + imp_rows, [55 * mm, 119 * mm]))
        story.append(Spacer(1, 10))

    # Nota calcoli
    disclaimer = data.get(
        "disclaimer_calcoli",
        "I calcoli e le verifiche riportate sono di sintesi e in linea con le normative applicabili"
        "Le raccomandazioni non sostituiscono le verifiche prescrittive previste dalle norme applicabili.",
    )
    if _meaningful(disclaimer):
        story.append(_p("NOTA", h2))
        story.append(_p(disclaimer, styles["BodyText"]))

    story.append(PageBreak())

    # === CAPITOLI 1..6 ===
    story.append(_p("CAPITOLO 1 - PREMESSA", h2))
    story.append(_p(data.get("premessa", ""), styles["BodyText"]))
    story.append(Spacer(1, 10))

    story.append(_p("CAPITOLO 2 - RIFERIMENTI LEGISLATIVI E NORMATIVI", h2))
    story.append(_p(data.get("norme", ""), styles["BodyText"]))
    story.append(Spacer(1, 10))

    criterio = data.get("criterio_progetto", "")
    if _meaningful(criterio):
        story.append(_p("CAPITOLO 3 - CRITERI DI PROGETTO DEGLI IMPIANTI", h2))
        story.append(_p(criterio, styles["BodyText"]))
        story.append(Spacer(1, 10))


    dati_proj = data.get("dati_progettazione", "")
    if _meaningful(dati_proj):
        story.append(_p("3.0 Dati raccolti in fase di progettazione (rilievo e presupposti)", h3))
        story.append(_p(dati_proj, styles["BodyText"]))
        story.append(Spacer(1, 10))

    story.append(_p("CAPITOLO 4 - SOLUZIONE PROGETTUALE ADOTTATA", h2))

    loc = data.get("localizzazione", "")
    if _meaningful(loc):
        story.append(_p("4.0 Localizzazione dell'impianto", h3))
        story.append(_p(loc, styles["BodyText"]))
        story.append(Spacer(1, 8))

    lay = data.get("layout", "")
    if _meaningful(lay):
        story.append(_p("4.0.1 Layout d'impianto", h3))
        story.append(_p(lay, styles["BodyText"]))
        story.append(Spacer(1, 8))

    dati_tecnici = data.get("dati_tecnici", "")
    if _meaningful(dati_tecnici):
        story.append(_p("4.1 Dati tecnici di base", h3))
        story.append(_p(dati_tecnici, styles["BodyText"]))
        story.append(Spacer(1, 8))

    descr = data.get("descrizione_impianto", "")
    if _meaningful(descr):
        story.append(_p("4.2 Descrizione impianto e opere", h3))
        story.append(_p(descr, styles["BodyText"]))
        story.append(Spacer(1, 8))

    ev_tbl = _evse_table(data, styles)
    if ev_tbl is not None:
        story.append(_p("4.2.1 Infrastruttura di ricarica EV (sintesi)", h3))
        story.append(ev_tbl)
        story.append(Spacer(1, 10))

    conf = data.get("confini", "")
    if _meaningful(conf):
        story.append(_p("4.3 Confini dell’intervento e interfacce", h3))
        story.append(_p(conf, styles["BodyText"]))
        story.append(Spacer(1, 10))

    quadri = data.get("quadri", [])
    if quadri:
        story.append(_p("4.4 Quadri elettrici e distribuzione (sintesi)", h3))
        tdata = [[
            _p("Quadro", th),
            _p("Ubicazione", th),
            _p("IP", th),
            _p("Interruttore generale<br/>(tipo/In)", th),
            _p("Differenziale generale<br/>(tipo/Idn)", th),
        ]]
        for q in quadri:
            tdata.append([
                _p(str(q.get("Quadro", "")), tc),
                _p(str(q.get("Ubicazione", "")), tc),
                _p(str(q.get("IP", "")), tc),
                _p(str(q.get("Generale", "")), tc),
                _p(str(q.get("Diff", "")), tc),
            ])
        tbl = Table(tdata, colWidths=[16 * mm, 40 * mm, 12 * mm, 52 * mm, 54 * mm], repeatRows=1, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 10))

    linee = data.get("linee", [])
    if linee:
        story.append(_p("4.5 Elenco circuiti, cavi e protezioni (sintesi)", h3))
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
            posa = (ln.get("Posa", "") or "").strip()
            ll = ln.get("L_m", "")
            posa_len = f"{posa}\n{ll}" if _meaningful(posa) else f"{ll}"
            tdata.append([
                _p(str(ln.get("Linea", "")), tc),
                _p(str(ln.get("Uso", "")), tc),
                _p(str(posa_len), tc),
                _p(str(ln.get("Cavo", "")), tc),
                _p(str(ln.get("Protezione", "")), tc),
                _p(str(ln.get("Diff", "")), tc),
                _p(str(ln.get("DV_perc", "")), tc),
                _p(str(ln.get("Esito", "")), tc),
            ])
        colw = [16 * mm, 30 * mm, 24 * mm, 32 * mm, 26 * mm, 26 * mm, 10 * mm, 10 * mm]
        tbl = Table(tdata, colWidths=colw, repeatRows=1, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 10))

    story.append(_p("CAPITOLO 5 - ULTERIORI INDICAZIONI", h2))

    sic = data.get("sicurezza", "")
    if _meaningful(sic):
        story.append(_p("5.1 Protezione contro i contatti diretti e indiretti", h3))
        story.append(_p(sic, styles["BodyText"]))
        story.append(Spacer(1, 8))

    ver = data.get("verifiche", "")
    ver_tab = data.get("verifiche_tabella") or []
    if _meaningful(ver) or ver_tab:
        story.append(_p("5.2 Verifiche, prove e collaudi", h3))
        if _meaningful(ver):
            story.append(_p(ver, styles["BodyText"]))
            story.append(Spacer(1, 6))
        if ver_tab:
            tdata = [[
                _p("Prova / Verifica", th),
                _p("Esito", th),
                _p("Strumento", th),
                _p("Note", th),
            ]]
            for r in ver_tab:
                tdata.append([
                    _p(str(r.get("Prova / Verifica","")), tc),
                    _p(str(r.get("Esito","")), tc),
                    _p(str(r.get("Strumento","")), tc),
                    _p(str(r.get("Note","")), tc),
                ])
            tbl = Table(tdata, colWidths=[60*mm, 20*mm, 35*mm, 55*mm], repeatRows=1, hAlign="LEFT")
            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]))
            story.append(tbl)
        story.append(Spacer(1, 8))

    man = data.get("manutenzione", "")
    if _meaningful(man):
        story.append(_p("5.3 Esercizio, manutenzione e avvertenze", h3))
        story.append(_p(man, styles["BodyText"]))
        story.append(Spacer(1, 8))

    allg = data.get("allegati", "")
    checklist = data.get("checklist_documentale") or []
    photo_flows = _photo_flowables(data)
    if _meaningful(allg) or checklist or photo_flows:
        story.append(_p("CAPITOLO 6 - ALLEGATI E DOCUMENTAZIONE", h2))

        if checklist:
            story.append(_p("6.0 Checklist documentale (D.M. 37/08, D.P.R. 462/2001 se applicabile)", h3))
            tdata = [[
                _p("Documento / Elaborato", th),
                _p("Stato", th),
                _p("Note", th),
            ]]
            for r in checklist:
                tdata.append([
                    _p(str(r.get("Documento / Elaborato","")), tc),
                    _p(str(r.get("Stato","")), tc),
                    _p(str(r.get("Note","")), tc),
                ])
            tbl = Table(tdata, colWidths=[95*mm, 25*mm, 50*mm], repeatRows=1, hAlign="LEFT")
            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]))
            story.append(tbl)
            story.append(Spacer(1, 8))

        if _meaningful(allg):
            story.append(_p(allg, styles["BodyText"]))
        if photo_flows:
            story.append(Spacer(1, 8))
            story.append(_p("6.1 Documentazione fotografica (estratto)", h3))
            story.extend(photo_flows)


    # CAPITOLO 7 - CARATTERISTICHE TECNICHE (estese, opzionali)
    tech_items = data.get("caratteristiche_tecniche_items") or []
    tech_note = data.get("caratteristiche_tecniche_note", "")
    if tech_items or _meaningful(tech_note):
        story.append(PageBreak())
        story.append(_p("CAPITOLO 7 - CARATTERISTICHE TECNICHE", h2))

        # Stampa i soli elementi compilati (nessun placeholder)
        for item in tech_items:
            titolo = str(item.get("titolo", "")).strip()
            testo = str(item.get("testo", "")).strip()
            if _meaningful(titolo) and _meaningful(testo):
                story.append(_p(titolo, h3))
                story.append(_p(testo, styles["BodyText"]))
                story.append(Spacer(1, 6))

        if _meaningful(tech_note):
            story.append(_p("Note / integrazioni", h3))
            story.append(_p(tech_note, styles["BodyText"]))

    # CAPITOLO 8 - ASPETTI NORMATIVI (estesi, opzionali)
    norm_items = data.get("aspetti_normativi_items") or []
    norm_note = data.get("aspetti_normativi_note", "")
    if norm_items or _meaningful(norm_note):
        story.append(PageBreak())
        story.append(_p("CAPITOLO 8 - ASPETTI NORMATIVI", h2))

        for item in norm_items:
            titolo = str(item.get("titolo", "")).strip()
            testo = str(item.get("testo", "")).strip()
            if _meaningful(titolo) and _meaningful(testo):
                story.append(_p(titolo, h3))
                story.append(_p(testo, styles["BodyText"]))
                story.append(Spacer(1, 6))

        if _meaningful(norm_note):
            story.append(_p("Note / integrazioni", h3))
            story.append(_p(norm_note, styles["BodyText"]))

    # Firma finale (facoltativa)
    luogo_f = data.get("luogo_firma", "")
    data_f = data.get("data_firma", "")
    firma = data.get("firma", "")
    if _meaningful(luogo_f) or _meaningful(data_f) or _meaningful(firma):
        story.append(Spacer(1, 14))
        if _meaningful(luogo_f) or _meaningful(data_f):
            story.append(_p(f"Luogo e data: {luogo_f} – {data_f}".strip(" –"), styles["BodyText"]))
            story.append(Spacer(1, 8))
        if _meaningful(firma):
            story.append(_p(f"Firma e timbro: {firma}", styles["BodyText"]))

    doc.build(
        story,
        onFirstPage=lambda c, d: None,
        onLaterPages=lambda c, d: _draw_header_footer(c, d, data),
        canvasmaker=_NumberedCanvas,
    )
    return buf.getvalue()
