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
    if "xxxx" in low:
        return False
    return True


def _first_non_empty(values: List[Any], default: str = "") -> str:
    """Return first meaningful value from list, else default.

    Keeps the document professional while ensuring that unanswered fields
    do not leak placeholders into the final PDF.
    """
    for v in values:
        if _meaningful(v):
            return str(v).strip()
    return default



def _std_premessa(data: Dict[str, Any]) -> str:
    comm = _first_non_empty([data.get("committente_nome"), "Committenza"])
    ogg = _first_non_empty([data.get("oggetto_intervento"), "l’intervento in oggetto"])
    luogo = _first_non_empty([data.get("impianto_indirizzo"), data.get("luogo"), "sito di intervento"])
    data_conf = _first_non_empty([data.get("data_conferma"), data.get("data_doc"), ""])
    fonte = _first_non_empty([data.get("fonte_dati"), "Committente"])
    # testo coerente con un report DiCo: tecnico + perimetro + responsabilità dati
    return (
        f"La presente Relazione Tecnico-Specialistica è redatta nell’ambito dell’incarico conferito dalla {comm} "
        f"e riguarda {ogg} presso {luogo}.\n\n"
        "<b>FINALITÀ E PERIMETRO</b>\n"
        "Il documento ha lo scopo di:\n"
        "• descrivere l’impianto e le opere eseguite/da eseguire, con indicazione dei confini dell’intervento;\n"
        "• richiamare i riferimenti legislativi e normativi applicabili;\n"
        "• esplicitare i criteri di progettazione e le verifiche di coordinamento essenziali (correnti, cadute di tensione, protezioni), "
        "in coerenza con la regola dell’arte.\n\n"
        "<b>VALENZA DOCUMENTALE</b>\n"
        "La presente Relazione costituisce documento tecnico di progetto e di supporto alla documentazione di conformità ai sensi del D.M. 37/2008; "
        "non sostituisce la Dichiarazione di Conformità (DiCo) né i relativi allegati obbligatori, che restano di competenza dell’Impresa installatrice.\n\n"
        "<b>RESPONSABILITÀ E DATI DI INGRESSO</b>\n"
        f"Le informazioni relative alla fornitura elettrica (POD, potenza disponibile/contrattuale, caratteristiche del punto di consegna), "
        f"destinazione d’uso e condizioni di esercizio sono state fornite da {fonte} e/o rilevate in sito"
        + (f" e/o confermate in data {data_conf}." if _meaningful(data_conf) else ".")
        + " Eventuali porzioni preesistenti non oggetto di intervento e le interfacce con impianti/parti terze sono indicate nel paragrafo “Confini dell’intervento”.\n\n"
        "<b>REQUISITI MATERIALI E CONSEGNA</b>\n"
        "Materiali e componenti devono essere conformi alle norme applicabili, provvisti di marcatura CE e, ove disponibile, marchio di conformità volontario "
        "(es. IMQ) o equivalente. Alla consegna l’impianto deve risultare conforme alla regola dell’arte e alle prescrizioni eventualmente impartite da Enti/Autorità competenti."
    )

def _std_norme(data: Dict[str, Any]) -> str:
    enti = data.get("prescrizioni_enti")
    enti_txt = enti if _meaningful(enti) else "Nessuna / Non applicabile"
    return (
        "Si riportano i principali riferimenti legislativi e normativi applicabili (elenco non esaustivo):\n"
        "• D.M. 22/01/2008 n. 37.\n"
        "• Legge 01/03/1968 n. 186.\n"
        "• D.Lgs. 09/04/2008 n. 81 e s.m.i.\n"
        "• D.P.R. 22/10/2001 n. 462 (ove applicabile).\n"
        "• Norme CEI applicabili (in particolare CEI 64-8, CEI 64-14, CEI EN 61439, CEI EN 60529; e, se pertinenti, CEI 81-10, CEI 0-10, CEI 0-21/0-16).\n"
        "• Regolamento Prodotti da Costruzione (UE) 305/2011 (CPR) e norme CEI-UNEL per i cavi (ove applicabile).\n"
        f"Eventuali ulteriori prescrizioni di Enti/Autorità locali: {enti_txt}."
    )

def _std_criteri_progetto(data: Dict[str, Any]) -> str:
    # Testo tecnico completo, coerente con relazione DiCo. I campi specifici (tipo cavo, posa, ecc.) sono riportati nelle tabelle di sintesi.
    return (
        "Tutti i materiali e le apparecchiature utilizzati devono essere di alta qualità, prodotti da aziende affidabili, ben lavorati e adatti all'uso previsto, "
        "resistendo a sollecitazioni meccaniche, corrosione, calore, umidità e acque meteoriche (per installazione all’esterno). Devono garantire lunga durata, "
        "facilità di ispezione e manutenzione.\n\n"
        "È obbligatorio l'uso di componenti con marcatura CE e, se disponibile, marchio IMQ o equivalente europeo. I componenti senza marcatura CE devono avere "
        "una dichiarazione di conformità del costruttore ai requisiti di sicurezza delle normative CEI, UNI o IEC.\n\n"
        "<b>3.1 Dimensionamento delle linee</b>\n"
        "Le linee elettriche sono calcolate mediante l’utilizzo dei seguenti criteri progettuali:\n"
        "• La corrente di impiego (Ib) è calcolata considerando la potenza nominale delle apparecchiature elettriche.\n"
        "• La corrente nominale della protezione (In) è considerata come la corrente che l’interruttore può sopportare per un tempo indefinito senza danni.\n"
        "• La portata del cavo (Iz) è valutata in funzione delle condizioni di posa e delle tabelle applicabili.\n\n"
        "<b>3.2 Sezione cavo in funzione di Ib</b>\n"
        "Nota la potenza assorbita dall’utenza, la corrente d’impiego (Ib) può essere calcolata come:\n"
        "Ib = (Ku · P) / (k · Vn · cosφ)\n"
        "dove k = 1 (monofase) o k = √3 (trifase). Determinata Ib, si dimensiona il cavo con portata Iz > Ib.\n\n"
        "<b>3.3 Caduta di tensione</b>\n"
        "La caduta di tensione percentuale complessiva non deve superare 4.0% (rif. CEI 64-8 art. 525), salvo diverse esigenze di progetto.\n\n"
        "<b>3.4 Cavi, posa e identificazione</b>\n"
        "I cavi utilizzati sono conformi al Regolamento UE 305/2011 (CPR), alle norme costruttive CEI e all’unificazione UNEL. "
        "I conduttori sono identificati secondo CEI-UNEL 00722 e 00712 (PE giallo/verde; neutro blu; fasi marrone/nero/grigio). "
        "Le modalità di posa e gli attraversamenti di pareti/solai devono mantenere, ove necessario, le prestazioni richieste (es. compartimentazioni).\n\n"
        "<b>3.5 Protezioni</b>\n"
        "La protezione dalle sovracorrenti è assicurata da interruttori automatici dimensionati e coordinati con le linee. "
        "Devono interrompere sovraccarichi e cortocircuiti prima di danni all’isolamento e avere potere di interruzione adeguato (PdI/Icu > Icc presunta).\n\n"
        "<b>3.6 Contatti indiretti</b>\n"
        "La protezione contro i contatti indiretti è realizzata mediante interruzione automatica dell’alimentazione (TT/TN) e/o componenti a doppio isolamento. "
        "Per sistemi TT si verifica il coordinamento Idn ≤ UL / Rt (UL = 50 V in ambienti ordinari).\n\n"
        "<b>3.7 Contatti diretti</b>\n"
        "La protezione contro i contatti diretti è assicurata tramite isolamento delle parti attive e/o involucri/barriere con grado di protezione adeguato (minimo IPXXB).\n\n"
        "<b>3.8 Potere di interruzione</b>\n"
        "Icc-max < PdI (Icu) del dispositivo di protezione (rif. CEI EN 60947-2).\n\n"
        "<b>3.9 Quadri elettrici</b>\n"
        "Quadri conformi a CEI EN 61439-1/2 (e/o CEI 23-51 per domestici/similari). Cablaggio interno con conduttori idonei e dimensionato per corrente nominale e cortocircuito nel punto di installazione."
    )
def _std_manutenzione() -> str:
    return (
        "Le attività di esercizio e manutenzione devono essere svolte da personale qualificato e autorizzato, in sicurezza e nel rispetto delle istruzioni "
        "dei costruttori e delle norme tecniche applicabili (es. CEI 0-10 / CEI 11-27, ove pertinenti).\n\n"
        "<b>PIANO DI MANUTENZIONE (minimo consigliato)</b>\n"
        "• Quadri elettrici: ispezione visiva, pulizia, verifica serraggi morsetti, integrità targhe/etichette e dispositivi di protezione;\n"
        "• Dispositivi differenziali: prova periodica con tasto “T” e verifiche strumentali (Idn/tempo) secondo periodicità e criticità del sito;\n"
        "• Conduttori e condutture: verifica integrità isolamento, fissaggi, protezioni meccaniche e segregazioni;\n"
        "• Collegamenti equipotenziali e PE: controllo continuità e integrità;\n"
        "• Comandi/emergenze (se presenti): prova funzionale e ripristino, verifica segnalazioni e cartellonistica;\n"
        "• Apparecchiature specifiche (es. wallbox/utenze dedicate): ispezione cavi e connettori, prova funzionale e aggiornamenti firmware se previsti dal costruttore.\n\n"
        "È raccomandata la tenuta di un registro manutenzione con data, attività eseguite, esito e nominativo dell’operatore."
    )

def _std_dati_tecnici_base(data: Dict[str, Any]) -> str:
    sist = _first_non_empty([data.get("sistema_distribuzione"), "TT"])
    tf = _first_non_empty([data.get("tensione_freq"), "230/400 V - 50 Hz"])
    pot = data.get("potenza_disponibile")
    pod = data.get("pod")
    cont = data.get("contatore_ubicazione")
    alim = data.get("alimentazione")
    amb = data.get("ambienti") or []
    amb_txt = ", ".join([a for a in amb if _meaningful(a)])
    parts = []
    parts.append(f"Tipo sistema di distribuzione: {sist}.")
    parts.append(f"Tensione nominale: {tf}.")
    if _meaningful(pot):
        parts.append(f"Potenza disponibile/contrattuale: {pot}.")
    if _meaningful(pod):
        parts.append(f"POD: {pod}.")
    if _meaningful(cont):
        parts.append(f"Contatore ubicato in: {cont}.")
    if _meaningful(alim):
        parts.append(f"Alimentazione: {alim}.")
    if _meaningful(amb_txt):
        parts.append(f"Ambientazioni particolari (se presenti): {amb_txt}.")
    return " ".join(parts)

def _std_descrizione_impianto(data: Dict[str, Any]) -> str:
    luogo = _first_non_empty([data.get("impianto_indirizzo"), data.get("luogo"), ""])
    pod = data.get("pod")
    cont = data.get("contatore_ubicazione")
    sist = _first_non_empty([data.get("sistema_distribuzione"), "TT"])
    tf = _first_non_empty([data.get("tensione_freq"), "230/400 V - 50 Hz"])
    pot = data.get("potenza_disponibile")
    s = []
    if _meaningful(luogo):
        s.append(f"Il sito di intervento è ubicato in {luogo}.")
    if _meaningful(pod) or _meaningful(cont):
        s.append("L’impianto è alimentato in bassa tensione dal punto di consegna del Distributore"
                 + (f" (POD: {pod})" if _meaningful(pod) else "")
                 + (f", tramite contatore/quadretto di misura ubicato in {cont}" if _meaningful(cont) else "")
                 + ".")
    s.append(f"Tipo sistema di distribuzione: {sist}. Tensione nominale: {tf}."
             + (f" Potenza disponibile/contrattuale: {pot}." if _meaningful(pot) else ""))
    s.append("La ripartizione e distribuzione interna avviene mediante linee in cavo conforme CEI/UNEL e componenti marcati CE "
             "(e, ove disponibile, IMQ o equivalente). Le condutture sono posate in tubazioni/canalizzazioni idonee e con protezione meccanica adeguata; "
             "i circuiti risultano identificati e separati per destinazione d’uso, privilegiando la manutenibilità.")
    s.append("Le opere impiantistiche previste comprendono, in funzione dell’intervento, la realizzazione e/o modifica di linee di alimentazione dedicate, "
             "installazione di punti di utilizzo, posa di tubazioni/canalizzazioni, installazione o adeguamento di quadri elettrici, apparecchi di protezione e comando, "
             "morsetterie e accessori, nonché collegamenti al sistema di protezione (PE) e ai collegamenti equipotenziali.")
    s.append("I conduttori sono identificati secondo codifica colori (PE giallo-verde, N blu, fasi marrone/nero/grigio) e marcatura/etichettatura dove previsto. "
             "I dispositivi di protezione sono coordinati con le linee e con il sistema di distribuzione (TT/TN) in modo coerente con le norme tecniche applicabili.")
    return " ".join(s)



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
    """Canvas che consente 'Pagina X di Y' senza duplicare le pagine (replay a fine build)."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states: List[dict] = []

    def showPage(self):
        # Salva lo stato della pagina corrente e avvia una nuova pagina
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        # Replay delle pagine salvate con numerazione completa
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
    # Struttura capitoli "editoriale" (come nel template campione)
    return [
        "CAPITOLO 1: Premessa",
        "CAPITOLO 2: Riferimenti Legislativi e normativi",
        "CAPITOLO 3: Criteri di progetto degli impianti",
        "CAPITOLO 4: Soluzione progettuale adottata",
        "CAPITOLO 5: Ulteriori indicazioni",
        "CAPITOLO 6: Allegati",
    ]



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
    story.append(_p(_std_premessa(data), styles["BodyText"]))
    story.append(Spacer(1, 10))

    story.append(_p("CAPITOLO 2 - RIFERIMENTI LEGISLATIVI E NORMATIVI", h2))
    story.append(_p(_std_norme(data), styles["BodyText"]))
    story.append(Spacer(1, 10))
    story.append(_p("CAPITOLO 3 - CRITERI DI PROGETTO DEGLI IMPIANTI", h2))
    story.append(_p(_std_criteri_progetto(data), styles["BodyText"]))
    story.append(Spacer(1, 10))

    story.append(_p("CAPITOLO 4 - SOLUZIONE PROGETTUALE ADOTTATA", h2))
    story.append(_p("4.1 Dati tecnici di base", h3))
    story.append(_p(_std_dati_tecnici_base(data), styles["BodyText"]))
    story.append(Spacer(1, 8))
    story.append(_p("4.2 Descrizione impianto e opere", h3))
    story.append(_p(_std_descrizione_impianto(data), styles["BodyText"]))
    story.append(Spacer(1, 8))

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
            _p("Interruttore generale (tipo/In)", th),
            _p("Differenziale generale (tipo/Idn)", th),
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
            _p("Circuito/Linea", th),
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
    story.append(_p("5.2 Verifiche, prove e collaudi", h3))
    story.append(_p("Ad ultimazione dei lavori, l’impianto è sottoposto alle verifiche previste dalla CEI 64-8 (Parte 6) e dalla CEI 64-14, con esecuzione e registrazione delle prove strumentali pertinenti al sistema di distribuzione (TT/TN) e alla tipologia di impianto.", styles["BodyText"]))
    vt = data.get("verifiche_tabella") or []
    vt_rows = []
    if isinstance(vt, list):
        for r in vt:
            if not isinstance(r, dict):
                continue
            prova = str(r.get("Prova", "")).strip()
            esito = str(r.get("Esito", "")).strip()
            strum = str(r.get("Strumento", "")).strip()
            note = str(r.get("Note", "")).strip()
            if _meaningful(esito) or _meaningful(strum) or _meaningful(note):
                vt_rows.append([prova, esito, strum, note])
    if vt_rows:
        story.append(Spacer(1, 6))
        story.append(_kv_table([["Prova", "Esito", "Strumento", "Note"]] + vt_rows, [70*mm, 30*mm, 40*mm, 40*mm]))
    story.append(Spacer(1, 8))
    story.append(_p("5.3 Esercizio, manutenzione e avvertenze", h3))
    story.append(_p(_std_manutenzione(), styles["BodyText"]))
    story.append(Spacer(1, 8))
    # CAPITOLO 6 - ALLEGATI (sezione sempre in fondo; stampa solo contenuti presenti)
    story.append(_p("CAPITOLO 6 - ALLEGATI", h2))
    # 6.1 Checklist documentale
    ck = data.get("checklist") or data.get("checklist_documentale") or []
    ck_rows = []
    if isinstance(ck, list):
        for r in ck:
            if not isinstance(r, dict):
                continue
            docu = str(r.get("Documento/Elaborato", r.get("Documento", ""))).strip()
            stato = str(r.get("Stato", "")).strip()
            note = str(r.get("Note", "")).strip()
            if _meaningful(stato):
                ck_rows.append([docu, stato, note])
    if ck_rows:
        story.append(_p("6.1 Checklist documentale", h3))
        story.append(_kv_table([["Documento/Elaborato","Stato","Note"]] + ck_rows, [90*mm, 30*mm, 50*mm]))
        story.append(Spacer(1, 8))
    # 6.2 Documentazione fotografica essenziale
    fotos = data.get("foto") or []
    if isinstance(fotos, list) and len(fotos) > 0:
        story.append(_p("6.2 Documentazione fotografica essenziale", h3))
        story.append(_foto_gallery(fotos, styles))
        story.append(Spacer(1, 8))

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
