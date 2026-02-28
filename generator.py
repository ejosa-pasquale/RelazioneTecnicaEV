\
from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Union

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt


# -------------------- Data models --------------------
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
class ColonninaItem:
    descrizione: str
    quantita: int = 1


@dataclass
class PhotoItem:
    filename: str
    content: bytes
    caption: str = ""


@dataclass
class AllegatoItem:
    filename: str
    content: bytes
    kind: str  # "pdf" | "docx"


@dataclass
class RelazioneData:
    luogo_data: str
    committente: str
    sito_indirizzo: str
    sito_cap_citta: str
    oggetto: str

    distanza_m: int = 60
    potenza_impegnata_kw: float = 4.0
    ik_trifase_ka: float = 10.0
    ik_monofase_ka: float = 6.0
    cavo_lunghezza_m: int = 60
    cavo_tipo: str = "FG16OM16 3G6 0.6/1kV"

    layout_incluso: str = ""
    layout_escluso: str = ""


# -------------------- Template anchors --------------------
ANCHORS = {
    "DITTA_ESECUTRICE": "{{DITTA_ESECUTRICE}}",
    "LAYOUT": "{{LAYOUT_DESCRITTIVO}}",
    "COLONNINE": "{{COLONNINE}}",
    "FOTO": "{{FOTO_GALLERY}}",
    "DIAGRAMMA": "{{DIAGRAMMA_IMPIANTO}}",
    "ALLEGATI": "{{ALLEGATI_SCHEDA_TECNICA}}",
}

HEADING_HINTS = {
    "DITTA_ESECUTRICE": ["DITTA ESECUTRICE", "IMPRESA ESECUTRICE"],
    "LAYOUT": ["LAYOUT D'IMPIANTO", "LAYOUT D’IMPIANTO"],
    "COLONNINE": ["COLONNINE", "WALLBOX", "STAZIONI DI RICARICA"],
    "FOTO": ["DOCUMENTAZIONE FOTOGRAFICA", "FOTOGRAFICA"],
    "DIAGRAMMA": ["SCHEMA", "DIAGRAMMA", "LAYOUT"],
    "ALLEGATI": ["ALLEGATI", "SCHEDA TECNICA", "SCHEDE TECNICHE"],
}

FIELD_PLACEHOLDERS = {
    "{{LUOGO_DATA}}": lambda d: d.luogo_data,
    "{{COMMITTENTE}}": lambda d: d.committente,
    "{{SITO_INDIRIZZO}}": lambda d: d.sito_indirizzo,
    "{{SITO_CAP_CITTA}}": lambda d: d.sito_cap_citta,
    "{{OGGETTO}}": lambda d: d.oggetto,
    "{{DISTANZA_M}}": lambda d: str(d.distanza_m),
    "{{POTENZA_IMPEGNATA_KW}}": lambda d: f"{d.potenza_impegnata_kw:g}",
    "{{IK_TRIFASE_KA}}": lambda d: f"{d.ik_trifase_ka:g}",
    "{{IK_MONOFASE_KA}}": lambda d: f"{d.ik_monofase_ka:g}",
    "{{CAVO_LUNGHEZZA_M}}": lambda d: str(d.cavo_lunghezza_m),
    "{{CAVO_TIPO}}": lambda d: d.cavo_tipo,
}

_SAMPLE_REPLACEMENTS = {
    # Example strings -> placeholder tokens (helps retrofit existing templates)
    "Busnago, 06/02/2024": "{{LUOGO_DATA}}",
    "Nome": "{{COMMITTENTE}}",
    "via della SS. Annunziata 32/A": "{{SITO_INDIRIZZO}}",
    "55100 Lucca": "{{SITO_CAP_CITTA}}",
    "nuovo impianto di ricarica per veicolo elettrico": "{{OGGETTO}}",
}


# -------------------- Doc helpers --------------------
def _paragraph_text(p) -> str:
    return p.text or ""

def _is_heading_like(text: str) -> bool:
    t = (text or "").strip()
    # headings like "4 ..." or "4\t..." or "4. ..."
    return bool(re.match(r"^\d+(\.|)\s+", t))

def _delete_paragraph(p):
    p._element.getparent().remove(p._element)

def _wipe_paragraph(p):
    for r in p.runs:
        r.text = ""

def _replace_in_paragraph(p, mapping: Dict[str, str]):
    full = "".join(r.text for r in p.runs) if p.runs else (p.text or "")
    new = full
    for old, newv in mapping.items():
        if old:
            new = new.replace(old, newv)
    if new != full:
        if not p.runs:
            p.text = new
        else:
            p.runs[0].text = new
            for r in p.runs[1:]:
                r.text = ""

def _replace_everywhere(doc: Document, mapping: Dict[str, str]):
    for p in doc.paragraphs:
        _replace_in_paragraph(p, mapping)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph(p, mapping)
    # header/footer too
    for section in doc.sections:
        for hf in [section.header, section.footer]:
            for p in hf.paragraphs:
                _replace_in_paragraph(p, mapping)

def _find_first_paragraph_containing(doc: Document, needles: List[str]):
    needles_l = [n.lower() for n in needles]
    for p in doc.paragraphs:
        tx = (_paragraph_text(p)).lower()
        if any(n in tx for n in needles_l):
            return p
    return None

def ensure_anchors(doc: Document) -> List[str]:
    """
    Crea i punti di aggancio nel CORPO del testo.
    Se l'anchor non esiste già, la inserisce come paragrafo subito dopo un heading coerente.
    Ritorna la lista di anchor creati.
    """
    created = []
    existing_text = "\n".join(p.text or "" for p in doc.paragraphs)

    for key, marker in ANCHORS.items():
        if marker in existing_text:
            continue

        hints = HEADING_HINTS.get(key, [])
        anchor_after = _find_first_paragraph_containing(doc, hints) if hints else None

        if anchor_after is None:
            # fallback: append at end
            p = doc.add_paragraph(marker)
            created.append(marker)
            continue

        # insert new paragraph after anchor_after
        new_p = doc.add_paragraph(marker)
        anchor_after._p.addnext(new_p._p)
        created.append(marker)

    return created

def retrofit_field_placeholders(doc: Document) -> int:
    """Prova a trasformare alcune stringhe campione in placeholder, così i campi finiscono nel corpo."""
    count = 0
    for p in doc.paragraphs:
        before = p.text
        for old, ph in _SAMPLE_REPLACEMENTS.items():
            if old in (p.text or ""):
                _replace_in_paragraph(p, {old: ph})
        if p.text != before:
            count += 1
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    before = p.text
                    for old, ph in _SAMPLE_REPLACEMENTS.items():
                        if old in (p.text or ""):
                            _replace_in_paragraph(p, {old: ph})
                    if p.text != before:
                        count += 1
    return count

def build_field_mapping(data: RelazioneData) -> Dict[str, str]:
    return {k: fn(data) for k, fn in FIELD_PLACEHOLDERS.items()}

# -------------------- Section writers --------------------
def _clear_section_after_heading(doc: Document, heading: str):
    """Cancella i paragrafi dopo l'heading fino al prossimo heading numerato."""
    paragraphs = list(doc.paragraphs)
    title = None
    for p in paragraphs:
        if heading.lower() in (_paragraph_text(p)).lower():
            title = p
            break
    if not title:
        return None
    try:
        idx = paragraphs.index(title)
    except ValueError:
        return title
    to_delete = []
    for p2 in paragraphs[idx+1:]:
        if _is_heading_like(p2.text):
            break
        to_delete.append(p2)
    for p2 in reversed(to_delete):
        try:
            _delete_paragraph(p2)
        except Exception:
            pass
    return title

def _insert_bullets_after(paragraph, doc: Document, lines: List[str]):
    elm = paragraph._p
    for line in lines:
        if not line.strip():
            continue
        bp = doc.add_paragraph(line.strip(), style="List Bullet" if "List Bullet" in doc.styles else None)
        elm.addnext(bp._p)
        elm = bp._p

def write_layout(doc: Document, data: RelazioneData):
    # Prefer anchor
    marker = ANCHORS["LAYOUT"]
    p = _find_first_paragraph_containing(doc, [marker])
    if p:
        _wipe_paragraph(p)
        anchor = p
    else:
        anchor = _clear_section_after_heading(doc, "LAYOUT D'IMPIANTO") or _clear_section_after_heading(doc, "LAYOUT D’IMPIANTO")
        if not anchor:
            return
    # Remove any existing content if marker was inside body section
    # write included/excluded
    elm = anchor._p
    if data.layout_incluso.strip():
        lbl = doc.add_paragraph("Incluso:")
        lbl.runs[0].bold = True if lbl.runs else True
        elm.addnext(lbl._p); elm = lbl._p
        _insert_bullets_after(lbl, doc, data.layout_incluso.splitlines())
        # move elm to last inserted bullet by scanning next siblings isn't easy; ok to keep.
    if data.layout_escluso.strip():
        lbl2 = doc.add_paragraph("Escluso:")
        lbl2.runs[0].bold = True if lbl2.runs else True
        elm.addnext(lbl2._p); elm = lbl2._p
        _insert_bullets_after(lbl2, doc, data.layout_escluso.splitlines())

def write_colonnine(doc: Document, colonnine: List[ColonninaItem]):
    marker = ANCHORS["COLONNINE"]
    p = _find_first_paragraph_containing(doc, [marker])
    if not p:
        return
    _wipe_paragraph(p)
    lines = [f"n. {c.quantita} — {c.descrizione}" for c in colonnine]
    _insert_bullets_after(p, doc, lines)

def write_ditta_esecutrice(doc: Document, esecutrice: EsecutriceData):
    marker = ANCHORS["DITTA_ESECUTRICE"]
    p = _find_first_paragraph_containing(doc, [marker])
    if not p:
        return
    _wipe_paragraph(p)
    lines = []
    if esecutrice.nome.strip():
        lines.append(esecutrice.nome.strip())
    if esecutrice.indirizzo.strip():
        lines.append(esecutrice.indirizzo.strip())
    if esecutrice.piva.strip():
        lines.append(f"P.IVA: {esecutrice.piva.strip()}")
    if lines:
        _insert_bullets_after(p, doc, lines)

def write_foto(doc: Document, photos: List[PhotoItem]):
    marker = ANCHORS["FOTO"]
    p = _find_first_paragraph_containing(doc, [marker])
    if not p or not photos:
        return
    _wipe_paragraph(p)
    cols = 2
    rows = (len(photos)+cols-1)//cols
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
            run.add_picture(io.BytesIO(item.content), width=Inches(3.2))
            if item.caption.strip():
                cap = cell.add_paragraph(item.caption.strip())
                try:
                    cap.style = doc.styles["Caption"]
                except Exception:
                    pass
            idx += 1
    p._p.addnext(table._tbl)

def write_diagramma(doc: Document, diagram_bytes: Optional[bytes]):
    marker = ANCHORS["DIAGRAMMA"]
    p = _find_first_paragraph_containing(doc, [marker])
    if not p or not diagram_bytes:
        return
    _wipe_paragraph(p)
    p.add_run().add_picture(io.BytesIO(diagram_bytes), width=Inches(6.5))

def write_allegati(doc: Document, allegati: List[AllegatoItem]):
    marker = ANCHORS["ALLEGATI"]
    p = _find_first_paragraph_containing(doc, [marker])
    if not p or not allegati:
        return
    _wipe_paragraph(p)
    lines = [f"{a.filename}" for a in allegati]
    _insert_bullets_after(p, doc, lines)

# -------------------- Cover writer --------------------
def insert_cover(doc: Document, data: RelazioneData, progettista: ProgettistaData, esecutrice: Optional[EsecutriceData] = None):
    body = doc._body._element

    def add_par(text: str, size: int, bold: bool, align: str):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return p

    p_title = add_par("RELAZIONE TECNICA", 22, True, "center")
    add_par(data.oggetto.upper(), 14, True, "center")
    add_par("", 11, False, "center")
    add_par(f"Sito: {data.sito_indirizzo} — {data.sito_cap_citta}", 11, False, "center")
    add_par(f"Committente: {data.committente}", 11, False, "center")
    add_par(data.luogo_data, 11, False, "center")
    add_par("", 11, False, "center")

    p_proj = add_par("PROGETTISTA", 12, True, "left")
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

    if esecutrice and (esecutrice.nome.strip() or esecutrice.indirizzo.strip() or esecutrice.piva.strip()):
        p_ex = doc.add_paragraph()
        p_ex.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p_ex.add_run("DITTA ESECUTRICE")
        r.bold = True
        elm.addnext(p_ex._p); elm = p_ex._p
        for line in [esecutrice.nome.strip(), esecutrice.indirizzo.strip(), (f"P.IVA: {esecutrice.piva.strip()}" if esecutrice.piva.strip() else "")]:
            if not line:
                continue
            pp = doc.add_paragraph(line)
            pp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            elm.addnext(pp._p); elm = pp._p

    pb = doc.add_page_break()

    # move cover to top
    body_list = list(body)
    i0 = body_list.index(p_title._p)
    i1 = body_list.index(pb._p)
    elems = body_list[i0:i1+1]
    for e in elems:
        body.remove(e)
    for e in reversed(elems):
        body.insert(0, e)

# -------------------- Main API --------------------
def prepare_template(template: Union[bytes, Path]) -> bytes:
    """Crea i punti di aggancio nel corpo del testo e converte alcuni campi base in placeholder."""
    doc = Document(io.BytesIO(template) if isinstance(template, (bytes, bytearray)) else str(template))
    retrofit_field_placeholders(doc)
    ensure_anchors(doc)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def generate_document(
    template: Union[bytes, Path],
    data: RelazioneData,
    progettista: ProgettistaData,
    esecutrice: Optional[EsecutriceData],
    colonnine: List[ColonninaItem],
    photos: List[PhotoItem],
    diagram_bytes: Optional[bytes],
    allegati: List[AllegatoItem],
) -> bytes:
    doc = Document(io.BytesIO(template) if isinstance(template, (bytes, bytearray)) else str(template))

    # 1) create anchors (so generation works even if template wasn't prepared)
    ensure_anchors(doc)

    # 2) placeholder fields
    ph = build_field_mapping(data)
    _replace_everywhere(doc, ph)

    # 3) cover (optional but gives clean header info)
    insert_cover(doc, data, progettista, esecutrice)

    # 4) sections
    if esecutrice:
        write_ditta_esecutrice(doc, esecutrice)
    write_layout(doc, data)
    if colonnine:
        write_colonnine(doc, colonnine)
    if photos:
        write_foto(doc, photos)
    if diagram_bytes:
        write_diagramma(doc, diagram_bytes)
    if allegati:
        write_allegati(doc, allegati)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
