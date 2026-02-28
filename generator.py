\
from __future__ import annotations

import io
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Union

from docx import Document
from docx.shared import Inches


@dataclass
class PhotoItem:
    filename: str
    content: bytes
    caption: str = ""


@dataclass
class RelazioneData:
    luogo_data: str
    committente_nome: str
    sito_indirizzo: str
    sito_cap_citta: str

    richiedente_nome: str
    richiedente_indirizzo: str
    richiedente_cap_citta: str

    oggetto: str

    distanza_m: int = 60
    potenza_impegnata_kw: float = 4.0
    potenza_wallbox_kw: float = 7.4
    cavo_tipo: str = "FG16OM16 3G6 0.6/1kV"
    cavo_lunghezza_m: int = 60
    modello_wallbox: str = "Cupra Charger Connect monofase"

    ik_trifase_ka: float = 10.0
    ik_monofase_ka: float = 6.0


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
    # Basato su stringhe presenti nel template. Se una stringa non esiste, viene ignorata.
    mapping = {
        "Busnago, 06/02/2024": data.luogo_data,
        "Nome": data.committente_nome,
        "via della SS. Annunziata 32/A": data.sito_indirizzo,
        "55100 Lucca": data.sito_cap_citta,

        # nel template no-logo il richiedente è "Ing Pasquale Senese"
        "Ing Pasquale Senese": data.richiedente_nome if data.richiedente_nome else "Ing Pasquale Senese",

        "nuovo impianto di ricarica per veicolo elettrico": data.oggetto,

        "situata a Lucca (LU) in via della ss. Annunziata 32/A":
            f"situata a {data.sito_cap_citta} in {data.sito_indirizzo}",
        "dista circa 60 metri": f"dista circa {data.distanza_m} metri",

        "4 kW.": f"{data.potenza_impegnata_kw:g} kW.",
        "c.a. 7,4 kW": f"c.a. {data.potenza_wallbox_kw:g} kW",

        "pari a 10 kA": f"pari a {data.ik_trifase_ka:g} kA",
        "pari a 6 kA": f"pari a {data.ik_monofase_ka:g} kA",

        "WallBox Cupra Charger Connect monofase": data.modello_wallbox,
        "60 m di cavo FG16OM16 3G6 0.6/1kV": f"{data.cavo_lunghezza_m} m di cavo {data.cavo_tipo}",
    }
    return mapping


def _find_paragraph_with_marker(doc: Document, marker: str):
    for p in doc.paragraphs:
        if marker in p.text:
            return p
    return None


def _insert_diagram(doc: Document, marker: str, diagram_bytes: bytes, width_in: float = 6.5):
    p = _find_paragraph_with_marker(doc, marker)
    if not p or not diagram_bytes:
        return
    for r in p.runs:
        r.text = ""
    run = p.add_run()
    run.add_picture(io.BytesIO(diagram_bytes), width=Inches(width_in))


def _insert_gallery(doc: Document, marker: str, photos: List[PhotoItem], cols: int = 2):
    p = _find_paragraph_with_marker(doc, marker)
    if not p or not photos:
        return
    for r in p.runs:
        r.text = ""

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
            # immagine
            run = par.add_run()
            try:
                run.add_picture(io.BytesIO(item.content), width=Inches(3.2))
            except Exception:
                par.add_run(f"[Immagine non valida: {item.filename}]")
            # didascalia
            if item.caption.strip():
                cap_p = cell.add_paragraph(item.caption.strip())
                try:
                    cap_p.style = doc.styles["Caption"]
                except Exception:
                    pass
            idx += 1

    # posiziona tabella subito dopo il paragrafo marker (hack XML)
    p._p.addnext(table._tbl)


def generate_docx_bytes(
    template: Union[Path, bytes],
    data: RelazioneData,
    photos: Optional[List[PhotoItem]] = None,
    diagram_bytes: Optional[bytes] = None,
) -> bytes:
    if isinstance(template, (bytes, bytearray)):
        doc = Document(io.BytesIO(template))
    else:
        doc = Document(str(template))

    _replace_text_everywhere(doc, build_mapping(data))

    # Placeholder nel template:
    # {{DIAGRAMMA_IMPIANTO}}
    # {{FOTO_GALLERY}}
    if diagram_bytes:
        _insert_diagram(doc, "{{DIAGRAMMA_IMPIANTO}}", diagram_bytes)
    if photos:
        _insert_gallery(doc, "{{FOTO_GALLERY}}", photos)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
