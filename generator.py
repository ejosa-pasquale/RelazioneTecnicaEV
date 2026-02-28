\
from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Optional

from docx import Document


@dataclass
class RelazioneData:
    # Anagrafica
    luogo_data: str  # es. "Busnago, 06/02/2024"
    committente_nome: str
    sito_indirizzo: str
    sito_cap_citta: str

    richiedente_nome: str
    richiedente_indirizzo: str
    richiedente_cap_citta: str

    oggetto: str  # es. "nuovo impianto di ricarica per veicolo elettrico"

    # Dati tecnici (principali)
    distanza_m: int = 60
    potenza_impegnata_kw: float = 4.0
    potenza_wallbox_kw: float = 7.4
    cavo_tipo: str = "FG16OM16 3G6 0.6/1kV"
    cavo_lunghezza_m: int = 60
    modello_wallbox: str = "Cupra Charger Connect monofase"

    # Correnti di corto
    ik_trifase_ka: float = 10.0
    ik_monofase_ka: float = 6.0


def _replace_text_everywhere(doc: Document, mapping: Dict[str, str]) -> None:
    """Sostituzione testuale semplice su paragrafi e celle tabella.
    Nota: non copre eventuali textbox/forme non supportate da python-docx.
    """
    def repl_in_paragraph(p):
        full = "".join(run.text for run in p.runs)
        new = full
        for old, newv in mapping.items():
            new = new.replace(old, newv)
        if new != full:
            # ricostruisci: metti tutto nel primo run e svuota gli altri
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
    """Mappa basata sul documento di esempio fornito.
    Qui sostituiamo le stringhe campione presenti nel DOCX con i valori inseriti.
    """
    mapping = {
        # intestazione anagrafica (dalla versione di esempio)
        "Busnago, 06/02/2024": data.luogo_data,
        "Nome": data.committente_nome,
        "via della SS. Annunziata 32/A": data.sito_indirizzo,
        "55100 Lucca": data.sito_cap_citta,
        "Telebit Spa": data.richiedente_nome,
        "Via S. Rocco, 65": data.richiedente_indirizzo,
        "20874 Busnago": data.richiedente_cap_citta,
        "nuovo impianto di ricarica per veicolo elettrico": data.oggetto,

        # localizzazione (testo corpo)
        "situata a Lucca (LU) in via della ss. Annunziata 32/A": f"situata a {data.sito_cap_citta} in {data.sito_indirizzo}",
        "dista circa 60 metri": f"dista circa {data.distanza_m} metri",

        # potenze
        "4 kW.": f"{data.potenza_impegnata_kw:g} kW.",
        "c.a. 7,4 kW": f"c.a. {data.potenza_wallbox_kw:g} kW",

        # correnti di corto
        "pari a 10 kA": f"pari a {data.ik_trifase_ka:g} kA",
        "pari a 6 kA": f"pari a {data.ik_monofase_ka:g} kA",

        # layout impianto
        "WallBox Cupra Charger Connect monofase": data.modello_wallbox,
        "60 m di cavo FG16OM16 3G6 0.6/1kV": f"{data.cavo_lunghezza_m} m di cavo {data.cavo_tipo}",
    }
    return mapping


def generate_docx(
    template_path: Path,
    out_path: Path,
    data: RelazioneData,
) -> Path:
    doc = Document(str(template_path))
    mapping = build_mapping(data)
    _replace_text_everywhere(doc, mapping)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def maybe_convert_to_pdf(docx_path: Path, pdf_path: Path) -> Optional[Path]:
    """Conversione opzionale a PDF via LibreOffice (se disponibile)."""
    import subprocess
    try:
        # usa 'soffice' se presente
        subprocess.run(["soffice", "--version"], capture_output=True, check=True)
    except Exception:
        return None

    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    # LibreOffice salva in output dir con stesso nome file
    subprocess.run([
        "soffice",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(pdf_path.parent),
        str(docx_path),
    ], check=True)

    produced = pdf_path.parent / (docx_path.stem + ".pdf")
    if produced.exists():
        if produced != pdf_path:
            produced.replace(pdf_path)
        return pdf_path
    return None
