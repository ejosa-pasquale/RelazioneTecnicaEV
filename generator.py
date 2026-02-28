\
from __future__ import annotations

import io
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Union

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
    descrizione: str
    quantita: int = 1


@dataclass
class AllegatoItem:
    filename: str
    content: bytes
    kind: str  # "docx" | "pdf"


@dataclass
class ProgettistaData:
    nome: str
    indirizzo: str
    cell: str
    email: str
    piva: str


@dataclass
class EsecutriceData:
    nome: str
    indirizzo: str
    piva: str


@dataclass
class RelazioneData:
    luogo_data: str
    committente_nome: str
    sito_indirizzo: str
    sito_cap_citta: str
    oggetto: str

    distanza_m: int = 60
    potenza_impegnata_kw: float = 4.0
    potenza_wallbox_kw: float = 7.4
    cavo_tipo: str = "FG16OM16 3G6 0.6/1kV"
    cavo_lunghezza_m: int = 60

    ik_trifase_ka: float = 10.0
    ik_monofase_ka: float = 6.0

    layout_incluso: str = ""
    layout_escluso: str = ""


# ----------------- helpers -----------------
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


def _find_paragraph_contains(doc: Document, needle: str):
    n = needle.lower()
    for p in doc.paragraphs:
        if n in (p.text or "").lower():
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
    p.add_run().add_picture(io.BytesIO(diagram_bytes), width=Inches(width_in))


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
        try:
            p.style = doc.styles["Heading 3"]
        except Exception:
            pass

    last_elm = p._p
    for line in lines:
        if not line.strip():
            continue
        bp = doc.add_paragraph(line.strip(), style="List Bullet" if "List Bullet" in doc.styles else None)
        last_elm.addnext(bp._p)
        last_elm = bp._p


def _insert_layout_text(doc: Document, marker: str, included: str, excluded: str):
    """
    Inserisce il layout:
    - se trova il placeholder {{LAYOUT_DESCRITTIVO}} lo sostituisce lì.
    - altrimenti prova a inserirlo subito dopo il titolo 'LAYOUT D'IMPIANTO' (fallback).
    """
    p = _find_paragraph_with_marker(doc, marker)
    anchor = None
    wipe_anchor = True

    if p:
        anchor = p
    else:
        # fallback: cerca il titolo
        title = _find_paragraph_contains(doc, "LAYOUT D'IMPIANTO")
        if not title:
            title = _find_paragraph_contains(doc, "LAYOUT D’IMPIANTO")  # apostrofo tipografico
        if title:
            anchor = title
            wipe_anchor = False  # non cancellare il titolo

    if not anchor:
        return

    if wipe_anchor:
        _wipe_paragraph_text(anchor)
        anchor.add_run("Layout d’impianto (descrizione):").bold = True

    elm = anchor._p
    def add_label(lbl: str):
        p_lbl = doc.add_paragraph(lbl)
        if p_lbl.runs:
            p_lbl.runs[0].bold = True
        return p_lbl

    if included.strip():
        p_inc = add_label("Incluso:")
        elm.addnext(p_inc._p); elm = p_inc._p
        for line in included.splitlines():
            if line.strip():
                bp = doc.add_paragraph(line.strip(), style="List Bullet" if "List Bullet" in doc.styles else None)
                elm.addnext(bp._p); elm = bp._p

    if excluded.strip():
        p_exc = add_label("Escluso:")
        elm.addnext(p_exc._p); elm = p_exc._p
        for line in excluded.splitlines():
            if line.strip():
                bp = doc.add_paragraph(line.strip(), style="List Bullet" if "List Bullet" in doc.styles else None)
                elm.addnext(bp._p); elm = bp._p


def _insert_esecutrice_placeholder(doc: Document, marker: str, esecutrice: EsecutriceData):
    p = _find_paragraph_with_marker(doc, marker)
    if not p:
        return
    _wipe_paragraph_text(p)
    lines = []
    if esecutrice.nome.strip():
        lines.append(esecutrice.nome.strip())
    if esecutrice.indirizzo.strip():
        lines.append(esecutrice.indirizzo.strip())
    if esecutrice.piva.strip():
        lines.append(f"P.IVA: {esecutrice.piva.strip()}")

    if not lines:
        return

    p.add_run("Ditta esecutrice:").bold = True
    elm = p._p
    for line in lines:
        pp = doc.add_paragraph(line)
        pp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elm.addnext(pp._p)
        elm = pp._p


def _insert_allegati(doc: Document, marker: str, allegati: List[AllegatoItem]):
    p = _find_paragraph_with_marker(doc, marker)
    if not p or not allegati:
        return
    _wipe_paragraph_text(p)
    p.add_run("Allegati (schede tecniche):").bold = True
    elm = p._p
    for a in allegati:
        bp = doc.add_paragraph(f"- {a.filename}")
        elm.addnext(bp._p); elm = bp._p
        if a.kind == "docx":
            # best-effort: aggiunge il testo in appendice
            try:
                sub = Document(io.BytesIO(a.content))
                sep = doc.add_paragraph(f"--- Contenuto allegato: {a.filename} ---")
                if sep.runs:
                    sep.runs[0].italic = True
                elm.addnext(sep._p); elm = sep._p
                for sp in sub.paragraphs:
                    if sp.text.strip():
                        ap = doc.add_paragraph(sp.text)
                        elm.addnext(ap._p); elm = ap._p
            except Exception:
                pass


def _insert_cover(doc: Document, data: RelazioneData, progettista: ProgettistaData, esecutrice: Optional[EsecutriceData] = None):
    """
    Cover page robusta:
    usa doc.add_page_break() (evita AttributeError su pb_run.add_page_break()).
    """
    body = doc._body._element

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

    p_title = add_par("RELAZIONE TECNICA", size_pt=22, bold=True, align="center")
    add_par(data.oggetto.upper(), size_pt=14, bold=True, align="center")
    add_par("", 12)
    add_par(f"Sito: {data.sito_indirizzo} — {data.sito_cap_citta}", size_pt=11, align="center")
    add_par(f"Committente: {data.committente_nome}", size_pt=11, align="center")
    add_par(data.luogo_data, size_pt=11, align="center")
    add_par("", 12)

    # Progettista
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
        elm.addnext(pp._p); elm = pp._p

    # Ditta esecutrice (opzionale)
    if esecutrice and (esecutrice.nome.strip() or esecutrice.indirizzo.strip() or esecutrice.piva.strip()):
        p_exec = doc.add_paragraph()
        p_exec.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p_exec.add_run("DITTA ESECUTRICE")
        r.bold = True
        elm.addnext(p_exec._p); elm = p_exec._p
        for line in [
            esecutrice.nome.strip(),
            esecutrice.indirizzo.strip(),
            (f"P.IVA: {esecutrice.piva.strip()}" if esecutrice.piva.strip() else ""),
        ]:
            if not line:
                continue
            pp = doc.add_paragraph(line)
            pp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            elm.addnext(pp._p); elm = pp._p

    pb = doc.add_page_break()

    # sposta blocco cover in testa
    body_list = list(body)
    i0 = body_list.index(p_title._p)
    i1 = body_list.index(pb._p)
    elems = body_list[i0:i1+1]
    for e in elems:
        body.remove(e)
    for e in reversed(elems):
        body.insert(0, e)


# ----------------- main -----------------
def generate_docx_bytes(
    template: Union[Path, bytes],
    data: RelazioneData,
    progettista: ProgettistaData,
    esecutrice: Optional[EsecutriceData] = None,
    colonnine: Optional[List[ColonninaItem]] = None,
    photos: Optional[List[PhotoItem]] = None,
    diagram_bytes: Optional[bytes] = None,
    allegati: Optional[List[AllegatoItem]] = None,
) -> bytes:
    if isinstance(template, (bytes, bytearray)):
        doc = Document(io.BytesIO(template))
    else:
        doc = Document(str(template))

    _insert_cover(doc, data, progettista, esecutrice)

    _replace_text_everywhere(doc, build_mapping(data))

    if colonnine:
        lines = [f"n. {c.quantita} — {c.descrizione}" for c in colonnine]
        _insert_bullets(doc, "{{COLONNINE}}", lines, heading="Colonnine previste")

    # Layout: ora ha fallback anche senza placeholder
    _insert_layout_text(doc, "{{LAYOUT_DESCRITTIVO}}", data.layout_incluso, data.layout_escluso)

    if esecutrice:
        _insert_esecutrice_placeholder(doc, "{{DITTA_ESECUTRICE}}", esecutrice)

    if diagram_bytes:
        _insert_diagram(doc, "{{DIAGRAMMA_IMPIANTO}}", diagram_bytes)
    if photos:
        _insert_gallery(doc, "{{FOTO_GALLERY}}", photos)
    if allegati:
        _insert_allegati(doc, "{{ALLEGATI_SCHEDA_TECNICA}}", allegati)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
