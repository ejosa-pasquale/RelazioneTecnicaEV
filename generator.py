\
from __future__ import annotations

import io
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Union, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt


@dataclass
class PhotoItem:
    filename: str
    content: bytes
    caption: str = ""


@dataclass
class ColonninaItem:
    descrizione: str  # testo libero (modello, potenza, connettori, ecc.)
    quantita: int = 1


@dataclass
class AllegatoItem:
    filename: str
    content: bytes
    kind: str  # "docx" | "pdf" | "other"


@dataclass
class ProgettistaData:
    nome: str
    indirizzo: str
    cell: str
    email: str
    piva: str


@dataclass
class RelazioneData:
    # Anagrafica
    luogo_data: str
    committente_nome: str
    sito_indirizzo: str
    sito_cap_citta: str

    oggetto: str

    # Dati tecnici (principali)
    distanza_m: int = 60
    potenza_impegnata_kw: float = 4.0
    potenza_wallbox_kw: float = 7.4
    cavo_tipo: str = "FG16OM16 3G6 0.6/1kV"
    cavo_lunghezza_m: int = 60

    # Correnti di corto
    ik_trifase_ka: float = 10.0
    ik_monofase_ka: float = 6.0

    # Layout descrittivo
    layout_incluso: str = ""
    layout_escluso: str = ""


def _replace_text_everywhere(doc: Document, mapping: Dict[str, str]) -> None:
    def repl_in_paragraph(p):
        full = "".join(run.text for run in p.runs)
        new = full
        for old, newv in mapping.items():
            if old:
                new = new.replace(old, newv)
        if new != full:
            if not p.runs:
                p.add_run(new)
            else:
                p.runs[0].text = new
                for r in p.runs[1:]:
                    r.text = ""

    for p in doc.paragraphs:
        repl_in_paragraph(p)

    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    repl_in_paragraph(p)


def build_mapping(data: RelazioneData) -> Dict[str, str]:
    # Sostituzioni "sicure" basate sulle stringhe campione nel template.
    return {
        "Busnago, 06/02/2024": data.luogo_data,
        "Nome": data.committente_nome,
        "via della SS. Annunziata 32/A": data.sito_indirizzo,
        "55100 Lucca": data.sito_cap_citta,
        "nuovo impianto di ricarica per veicolo elettrico": data.oggetto,

        "situata a Lucca (LU) in via della ss. Annunziata 32/A":
            f"situata a {data.sito_cap_citta} in {data.sito_indirizzo}",
        "dista circa 60 metri": f"dista circa {data.distanza_m} metri",

        "4 kW.": f"{data.potenza_impegnata_kw:g} kW.",
        "c.a. 7,4 kW": f"c.a. {data.potenza_wallbox_kw:g} kW",

        "pari a 10 kA": f"pari a {data.ik_trifase_ka:g} kA",
        "pari a 6 kA": f"pari a {data.ik_monofase_ka:g} kA",

        "60 m di cavo FG16OM16 3G6 0.6/1kV": f"{data.cavo_lunghezza_m} m di cavo {data.cavo_tipo}",
    }


def _find_paragraph_with_marker(doc: Document, marker: str):
    for p in doc.paragraphs:
        if marker in p.text:
            return p
    return None


def _wipe_paragraph_text(p):
    for r in p.runs:
        r.text = ""


def _insert_diagram(doc: Document, marker: str, diagram_bytes: bytes, width_in: float = 6.5):
    p = _find_paragraph_with_marker(doc, marker)
    if not p or not diagram_bytes:
        return
    _wipe_paragraph_text(p)
    run = p.add_run()
    run.add_picture(io.BytesIO(diagram_bytes), width=Inches(width_in))


def _insert_gallery(doc: Document, marker: str, photos: List[PhotoItem], cols: int = 2):
    p = _find_paragraph_with_marker(doc, marker)
    if not p or not photos:
        return
    _wipe_paragraph_text(p)

    rows = (len(photos) + cols - 1) // cols
    table = doc.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"

    idx = 0
    for r in range(rows):
        for c in range(cols):
            cell = table.cell(r, c)
            if idx >= len(photos):
                cell.text = ""
                continue
            item = photos[idx]
            par = cell.paragraphs[0]
            run = par.add_run()
            try:
                run.add_picture(io.BytesIO(item.content), width=Inches(3.2))
            except Exception:
                par.add_run(f"[Immagine non valida: {item.filename}]")
            if item.caption.strip():
                cap_p = cell.add_paragraph(item.caption.strip())
                try:
                    cap_p.style = doc.styles["Caption"]
                except Exception:
                    pass
            idx += 1

    p._p.addnext(table._tbl)


def _insert_bullets(doc: Document, marker: str, lines: List[str], heading: Optional[str] = None):
    p = _find_paragraph_with_marker(doc, marker)
    if not p:
        return
    _wipe_paragraph_text(p)
    if heading:
        p.add_run(heading)
        p.style = doc.styles["Heading 3"] if "Heading 3" in doc.styles else p.style

    # Inserisci lista subito dopo
    last_elm = p._p
    for line in lines:
        if not line.strip():
            continue
        bp = doc.add_paragraph(line.strip(), style="List Bullet" if "List Bullet" in doc.styles else None)
        last_elm.addnext(bp._p)
        last_elm = bp._p


def _insert_layout_text(doc: Document, marker: str, included: str, excluded: str):
    p = _find_paragraph_with_marker(doc, marker)
    if not p:
        return
    _wipe_paragraph_text(p)
    p.add_run("Layout d’impianto (descrizione):").bold = True
    # paragrafi successivi
    elm = p._p
    if included.strip():
        p_inc = doc.add_paragraph("Incluso:")
        p_inc.runs[0].bold = True
        elm.addnext(p_inc._p); elm = p_inc._p
        for line in included.splitlines():
            if line.strip():
                bp = doc.add_paragraph(line.strip(), style="List Bullet" if "List Bullet" in doc.styles else None)
                elm.addnext(bp._p); elm = bp._p
    if excluded.strip():
        p_exc = doc.add_paragraph("Escluso:")
        p_exc.runs[0].bold = True
        elm.addnext(p_exc._p); elm = p_exc._p
        for line in excluded.splitlines():
            if line.strip():
                bp = doc.add_paragraph(line.strip(), style="List Bullet" if "List Bullet" in doc.styles else None)
                elm.addnext(bp._p); elm = bp._p


def _insert_cover(doc: Document, data: RelazioneData, progettista: ProgettistaData):
    # Inserisce una cover pulita all'inizio del documento (migliora impaginazione).
    # Nota: se il template contiene una cover "grafica" come forme Word, conviene eliminarla dal template.
    body = doc._body._element
    first = body[0]

    def add_par(text: str, size_pt: int = 12, bold: bool = False, align="center"):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size_pt)
        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "left":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return p

    # Costruisci in coda e poi spostiamo in testa (hack XML)
    p_title = add_par("RELAZIONE TECNICA", size_pt=22, bold=True, align="center")
    p_sub = add_par(data.oggetto.upper(), size_pt=14, bold=True, align="center")
    add_par("", 12)

    p_site = add_par(f"Sito: {data.sito_indirizzo} — {data.sito_cap_citta}", size_pt=11, align="center")
    p_comm = add_par(f"Committente: {data.committente_nome}", size_pt=11, align="center")
    p_date = add_par(data.luogo_data, size_pt=11, align="center")

    add_par("", 12)

    # blocco progettista a sinistra
    p_proj = add_par("PROGETTISTA", size_pt=12, bold=True, align="left")
    elm = p_proj._p
    for line in [
        progettista.nome,
        progettista.indirizzo,
        f"Cell: {progettista.cell}",
        f"Email: {progettista.email}",
        f"P.IVA: {progettista.piva}",
    ]:
        pp = doc.add_paragraph(line)
        pp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elm.addnext(pp._p)
        elm = pp._p

    # page break
    pb = doc.add_paragraph()
    pb_run = pb.add_run()
    pb_run.add_break()  # line break
    pb_run.add_break()  # keep
    pb_run.add_break()  # keep
    pb_run.add_break()  # keep
    pb_run.add_page_break()

    # Ora sposta tutti i paragrafi appena creati in testa (in ordine)
    # Raccogliamo gli ultimi N elementi aggiunti (da p_title fino a pb)
    new_elems = []
    # gli elementi sono stati aggiunti in fondo al body
    # prendiamo dall'ultimo page break paragraph fino al titolo (che è prima)
    # più semplice: individua riferimenti e sposta in ordine
    for p in [p_title, p_sub, p_site, p_comm, p_date, p_proj, pb]:
        new_elems.append(p._p)

    # In realtà abbiamo anche righe progettista create via add_paragraph in mezzo:
    # le troviamo tra p_proj e pb nel body: prendiamo tutti gli elementi dal p_title al pb
    # usando posizione in body
    body_list = list(body)
    i0 = body_list.index(p_title._p)
    i1 = body_list.index(pb._p)
    elems = body_list[i0:i1+1]
    # rimuovi e reinserisci in testa
    for e in elems:
        body.remove(e)
    for e in reversed(elems):
        body.insert(0, e)


def _insert_allegati(doc: Document, marker: str, allegati: List[AllegatoItem]):
    p = _find_paragraph_with_marker(doc, marker)
    if not p or not allegati:
        return
    _wipe_paragraph_text(p)
    p.add_run("Allegati (schede tecniche):").bold = True
    elm = p._p
    for a in allegati:
        line = f"- {a.filename}"
        bp = doc.add_paragraph(line)
        elm.addnext(bp._p)
        elm = bp._p

        # Se è DOCX, appendi contenuto come "Appendice"
        if a.kind == "docx":
            try:
                sub = Document(io.BytesIO(a.content))
                # separatore
                sep = doc.add_paragraph(f"--- Contenuto allegato: {a.filename} ---")
                sep.runs[0].italic = True
                elm.addnext(sep._p); elm = sep._p
                for sp in sub.paragraphs:
                    if sp.text.strip():
                        ap = doc.add_paragraph(sp.text)
                        elm.addnext(ap._p); elm = ap._p
            except Exception:
                pass


def generate_docx_bytes(
    template: Union[Path, bytes],
    data: RelazioneData,
    progettista: ProgettistaData,
    colonnine: Optional[List[ColonninaItem]] = None,
    photos: Optional[List[PhotoItem]] = None,
    diagram_bytes: Optional[bytes] = None,
    allegati: Optional[List[AllegatoItem]] = None,
) -> bytes:
    if isinstance(template, (bytes, bytearray)):
        doc = Document(io.BytesIO(template))
    else:
        doc = Document(str(template))

    # 1) Cover pulita (migliora la cover page)
    _insert_cover(doc, data, progettista)

    # 2) Sostituzioni base
    _replace_text_everywhere(doc, build_mapping(data))

    # 3) Layout descrittivo / colonnine / allegati / immagini
    if colonnine:
        lines = [f"n. {c.quantita} — {c.descrizione}" for c in colonnine]
        _insert_bullets(doc, "{{COLONNINE}}", lines, heading="Colonnine previste")
    _insert_layout_text(doc, "{{LAYOUT_DESCRITTIVO}}", data.layout_incluso, data.layout_escluso)

    if diagram_bytes:
        _insert_diagram(doc, "{{DIAGRAMMA_IMPIANTO}}", diagram_bytes)
    if photos:
        _insert_gallery(doc, "{{FOTO_GALLERY}}", photos)
    if allegati:
        _insert_allegati(doc, "{{ALLEGATI_SCHEDA_TECNICA}}", allegati)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
