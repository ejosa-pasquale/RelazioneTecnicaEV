\
import datetime as dt
import json
from pathlib import Path

import streamlit as st
from streamlit_drawable_canvas import st_canvas

from docx import Document
import io

def template_has_marker(template_bytes: bytes, marker: str) -> bool:
    try:
        d = Document(io.BytesIO(template_bytes))
        return any(marker in (p.text or "") for p in d.paragraphs)
    except Exception:
        return False


from generator import (
    PhotoItem, ColonninaItem, AllegatoItem,
    ProgettistaData, EsecutriceData, RelazioneData, generate_docx_bytes
)

APP_DIR = Path(__file__).parent
DEFAULT_TEMPLATE_PATH = APP_DIR / "templates" / "relazione_base.docx"
PROFILES_DIR = APP_DIR / "profiles"

st.set_page_config(page_title="Generatore Relazione Progetto Elettrico", layout="wide")
st.title("Generatore relazione tecnica (progetto elettrico)")
st.caption("Compila i campi, aggiungi foto/diagramma/allegati e scarica il DOCX generato.")


def load_profile(name: str) -> dict:
    p = PROFILES_DIR / f"{name}.json"
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return {}

def save_profile(name: str, data: dict) -> None:
    PROFILES_DIR.mkdir(parents=True, exist_ok=True)
    p = PROFILES_DIR / f"{name}.json"
    p.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


# ---------- Session state ----------
if "photos" not in st.session_state:
    st.session_state.photos = []
if "diagram_bytes" not in st.session_state:
    st.session_state.diagram_bytes = None
if "template_bytes" not in st.session_state:
    st.session_state.template_bytes = DEFAULT_TEMPLATE_PATH.read_bytes()
if "allegati" not in st.session_state:
    st.session_state.allegati = []
if "colonnine" not in st.session_state:
    st.session_state.colonnine = []


# ---------- Sidebar ----------
with st.sidebar:
    st.header("Profilo progetto (dati fissi)")
    existing_profiles = sorted([p.stem for p in PROFILES_DIR.glob("*.json")]) or ["default"]
    profile_name = st.selectbox("Seleziona profilo", options=existing_profiles, index=0)
    profile = load_profile(profile_name) or load_profile("default")

    st.divider()
    st.header("Progettista")
    progettista_nome = st.text_input("Nome", value=profile.get("progettista_nome", "Ing. Pasquale Senese"))
    progettista_indirizzo = st.text_input("Indirizzo", value=profile.get("progettista_indirizzo", "Via Francesco Soave 30 - 20135 Milano (MI)"))
    progettista_cell = st.text_input("Cell", value=profile.get("progettista_cell", "340 5731381"))
    progettista_email = st.text_input("Email", value=profile.get("progettista_email", "pasquale.senese@ingpec.eu"))
    progettista_piva = st.text_input("P.IVA", value=profile.get("progettista_piva", "14572980960"))

    st.divider()
    st.header("Ditta esecutrice")
    esecutrice_nome = st.text_input("Nome / Ragione sociale", value=profile.get("esecutrice_nome", ""))
    esecutrice_indirizzo = st.text_input("Indirizzo", value=profile.get("esecutrice_indirizzo", ""))
    esecutrice_piva = st.text_input("P.IVA", value=profile.get("esecutrice_piva", ""))

    st.divider()
    st.header("Dati anagrafici")
    luogo = st.text_input("Luogo (es. Busnago)", value=profile.get("luogo", "Busnago"))
    data = st.date_input("Data", value=dt.date.today())
    luogo_data = f"{luogo}, {data.strftime('%d/%m/%Y')}"

    committente_nome = st.text_input("Committente (Nome / Condomìnio)", value=profile.get("committente_nome", "Nome"))
    sito_indirizzo = st.text_input("Indirizzo sito", value=profile.get("sito_indirizzo", "via della SS. Annunziata 32/A"))
    sito_cap_citta = st.text_input("CAP e Città (sito)", value=profile.get("sito_cap_citta", "55100 Lucca"))

    st.header("Oggetto")
    oggetto = st.text_input("Oggetto intervento", value=profile.get("oggetto", "nuovo impianto di ricarica per veicolo elettrico"))

    st.header("Dati tecnici principali")
    distanza_m = st.number_input("Distanza punto origine (m)", min_value=1, value=int(profile.get("distanza_m", 60)), step=1)
    potenza_impegnata_kw = st.number_input("Potenza impegnata (kW)", min_value=0.5, value=float(profile.get("potenza_impegnata_kw", 4.0)), step=0.5)

    cavo_lunghezza_m = st.number_input("Lunghezza cavo (m)", min_value=1, value=int(profile.get("cavo_lunghezza_m", 60)), step=1)
    cavo_tipo = st.text_input("Tipo cavo", value=profile.get("cavo_tipo", "FG16OM16 3G6 0.6/1kV"))

    st.header("Corto circuito (punto fornitura)")
    ik_trifase_ka = st.number_input("Ik trifase presunta (kA)", min_value=0.1, value=float(profile.get("ik_trifase_ka", 10.0)), step=0.1)
    ik_monofase_ka = st.number_input("Ik monofase presunta (kA)", min_value=0.1, value=float(profile.get("ik_monofase_ka", 6.0)), step=0.1)

    st.divider()
    st.header("LAYOUT D'IMPIANTO (descrittivo)")
    layout_incluso = st.text_area("Cosa è incluso (una riga = un bullet)", value=profile.get("layout_incluso",""), height=120)
    layout_escluso = st.text_area("Cosa è escluso (una riga = un bullet)", value=profile.get("layout_escluso",""), height=120)

    st.divider()
    st.subheader("Salva profilo")
    new_profile_name = st.text_input("Nome profilo (per salvare)", value=profile_name)
    if st.button("Salva / aggiorna profilo", use_container_width=True):
        try:
            save_profile(new_profile_name, {
                **profile,
                "progettista_nome": progettista_nome,
                "progettista_indirizzo": progettista_indirizzo,
                "progettista_cell": progettista_cell,
                "progettista_email": progettista_email,
                "progettista_piva": progettista_piva,
                "esecutrice_nome": esecutrice_nome,
                "esecutrice_indirizzo": esecutrice_indirizzo,
                "esecutrice_piva": esecutrice_piva,
                "luogo": luogo,
                "committente_nome": committente_nome,
                "sito_indirizzo": sito_indirizzo,
                "sito_cap_citta": sito_cap_citta,
                "oggetto": oggetto,
                "distanza_m": distanza_m,
                "potenza_impegnata_kw": potenza_impegnata_kw,
                "cavo_lunghezza_m": cavo_lunghezza_m,
                "cavo_tipo": cavo_tipo,
                "ik_trifase_ka": ik_trifase_ka,
                "ik_monofase_ka": ik_monofase_ka,
                "layout_incluso": layout_incluso,
                "layout_escluso": layout_escluso,
            })
            st.success("Profilo salvato.")
        except Exception as e:
            st.warning("Non riesco a scrivere su disco (tipico su Streamlit Cloud). Profilo non salvato.")
            st.caption(str(e))

    st.divider()
    st.subheader("Template (senza logo Telebit)")
    uploaded_template = st.file_uploader("Sostituisci template (.docx)", type=["docx"])
    if uploaded_template:
        st.session_state.template_bytes = uploaded_template.getvalue()
        st.success("Template caricato (in memoria).")

    st.caption(
        "Placeholder nel template:\n"
        "- {{DITTA_ESECUTRICE}} (opzionale)\n"
        "- {{LAYOUT_DESCRITTIVO}}\n"
        "- {{COLONNINE}}\n"
        "- {{FOTO_GALLERY}}\n"
        "- {{DIAGRAMMA_IMPIANTO}}\n"
        "- {{ALLEGATI_SCHEDA_TECNICA}}"
    )


# ---------- Tabs ----------
tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["Generazione & Download", "Colonnine (multiplo)", "Foto", "Diagramma", "Schede tecniche (upload)"]
)

with tab2:
    st.subheader("Colonnine (multiplo)")
    c1, c2 = st.columns([3, 1])
    with c1:
        descr = st.text_input("Descrizione colonnina (modello, potenza, connettori, note)", value="")
    with c2:
        qty = st.number_input("Qtà", min_value=1, value=1, step=1)
    if st.button("Aggiungi colonnina", use_container_width=True):
        if descr.strip():
            st.session_state.colonnine.append(ColonninaItem(descrizione=descr.strip(), quantita=int(qty)))
            st.success("Colonnina aggiunta.")
        else:
            st.warning("Inserisci una descrizione.")

    if st.session_state.colonnine:
        st.divider()
        for i, item in enumerate(list(st.session_state.colonnine)):
            cc1, cc2, cc3 = st.columns([0.7, 3, 0.8])
            with cc1:
                st.write(f"n. {item.quantita}")
            with cc2:
                newd = st.text_input("Descrizione", value=item.descrizione, key=f"col_desc_{i}")
                st.session_state.colonnine[i].descrizione = newd
            with cc3:
                if st.button("Rimuovi", key=f"col_rm_{i}"):
                    st.session_state.colonnine.pop(i)
                    st.rerun()
    else:
        st.info("Nessuna colonnina inserita (opzionale).")

with tab5:
    st.subheader("Schede tecniche (upload)")
    ups = st.file_uploader("Carica schede tecniche (DOCX/PDF)", type=["pdf", "docx"], accept_multiple_files=True)
    if ups:
        existing = {a.filename for a in st.session_state.allegati}
        added = 0
        for u in ups:
            if u.name in existing:
                continue
            kind = "pdf" if u.name.lower().endswith(".pdf") else "docx"
            st.session_state.allegati.append(AllegatoItem(filename=u.name, content=u.getvalue(), kind=kind))
            added += 1
        if added:
            st.success(f"Aggiunti {added} allegati.")
    if st.session_state.allegati:
        for i, a in enumerate(list(st.session_state.allegati)):
            c1, c2 = st.columns([4, 1])
            with c1:
                st.write(f"📎 {a.filename} ({a.kind})")
            with c2:
                if st.button("Rimuovi", key=f"all_rm_{i}"):
                    st.session_state.allegati.pop(i)
                    st.rerun()
        st.caption("Nel template: {{ALLEGATI_SCHEDA_TECNICA}}")
    else:
        st.info("Nessun allegato caricato (opzionale).")

with tab3:
    st.subheader("Foto")
    uploads = st.file_uploader("Carica foto (JPG/PNG)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
    if uploads:
        existing = {p.filename for p in st.session_state.photos}
        for u in uploads:
            if u.name not in existing:
                st.session_state.photos.append(PhotoItem(filename=u.name, content=u.getvalue(), caption=""))
        st.success("Foto aggiunte.")

    if st.session_state.photos:
        for i, item in enumerate(list(st.session_state.photos)):
            c1, c2, c3 = st.columns([1, 2, 0.6])
            with c1:
                st.image(item.content, use_container_width=True)
            with c2:
                cap = st.text_input(f"Didascalia — {item.filename}", value=item.caption, key=f"cap_{i}")
                st.session_state.photos[i].caption = cap
            with c3:
                if st.button("Rimuovi", key=f"rm_{i}"):
                    st.session_state.photos.pop(i)
                    st.rerun()
        st.caption("Nel template: {{FOTO_GALLERY}}")
    else:
        st.info("Nessuna foto caricata (opzionale).")

with tab4:
    st.subheader("Diagramma impianto")
    mode = st.radio("Metodo", ["Carica immagine", "Disegna"], horizontal=True)

    if mode == "Carica immagine":
        diag_up = st.file_uploader("Carica diagramma (JPG/PNG)", type=["jpg", "jpeg", "png"], accept_multiple_files=False)
        if diag_up:
            st.session_state.diagram_bytes = diag_up.getvalue()
            st.image(st.session_state.diagram_bytes, caption="Diagramma selezionato", use_container_width=True)
    else:
        st.caption("Disegno semplice: linee/annotazioni rapide. Poi salva il disegno come diagramma.")
        canvas = st_canvas(
            fill_color="rgba(255, 255, 255, 0)",
            stroke_width=3,
            stroke_color="#000000",
            background_color="#FFFFFF",
            height=450,
            drawing_mode="freedraw",
            key="canvas",
        )
        if st.button("Usa disegno come diagramma", type="primary"):
            if canvas.image_data is not None:
                import PIL.Image
                import io
                img = PIL.Image.fromarray(canvas.image_data.astype("uint8"), mode="RGBA").convert("RGB")
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                st.session_state.diagram_bytes = buf.getvalue()
                st.success("Diagramma salvato.")
        if st.session_state.diagram_bytes:
            st.image(st.session_state.diagram_bytes, caption="Diagramma pronto", use_container_width=True)

    st.caption("Nel template: {{DIAGRAMMA_IMPIANTO}}")


with tab1:
    st.subheader("Generazione & Download")
    st.info("La cover include Progettista e (se compilata) la Ditta esecutrice.")

    # --- Sezioni del template (placeholder -> stato) ---
    sec = [
        ("Ditta esecutrice", "{{DITTA_ESECUTRICE}}", "Opzionale"),
        ("Layout d'impianto", "{{LAYOUT_DESCRITTIVO}}", "Richiesto se presente nel template"),
        ("Colonnine", "{{COLONNINE}}", "Richiesto se presente nel template"),
        ("Foto gallery", "{{FOTO_GALLERY}}", "Richiesto se presente nel template"),
        ("Diagramma impianto", "{{DIAGRAMMA_IMPIANTO}}", "Richiesto se presente nel template"),
        ("Allegati scheda tecnica", "{{ALLEGATI_SCHEDA_TECNICA}}", "Richiesto se presente nel template"),
    ]
    st.markdown("### Sezioni rilevate nel template")
    for label, marker, note in sec:
        present = template_has_marker(st.session_state.template_bytes, marker)
        st.write(("✅" if present else "➖") + f" **{label}** — {note}")


    
    missing = []
    for m in ["{{LAYOUT_DESCRITTIVO}}","{{COLONNINE}}","{{FOTO_GALLERY}}","{{DIAGRAMMA_IMPIANTO}}","{{ALLEGATI_SCHEDA_TECNICA}}","{{DITTA_ESECUTRICE}}"]:
        if not template_has_marker(st.session_state.template_bytes, m):
            missing.append(m)
    if missing:
        st.caption("Nel template mancano questi placeholder (alcune sezioni useranno fallback o non verranno inserite): " + ", ".join(missing))


    st.warning(
        "Placeholder nel template:\n"
        "- {{DITTA_ESECUTRICE}} (opzionale)\n"
        "- {{LAYOUT_DESCRITTIVO}}\n"
        "- {{COLONNINE}}\n"
        "- {{FOTO_GALLERY}}\n"
        "- {{DIAGRAMMA_IMPIANTO}}\n"
        "- {{ALLEGATI_SCHEDA_TECNICA}}"
    )

    filename_base = st.text_input("Nome file", value="relazione_progetto_elettrico")

    if st.button("Genera e prepara download", type="primary", use_container_width=True):
        # --- Requisiti sezioni (in base ai placeholder presenti nel template) ---
        present = {
            "{{DITTA_ESECUTRICE}}": template_has_marker(st.session_state.template_bytes, "{{DITTA_ESECUTRICE}}"),
            "{{LAYOUT_DESCRITTIVO}}": template_has_marker(st.session_state.template_bytes, "{{LAYOUT_DESCRITTIVO}}"),
            "{{COLONNINE}}": template_has_marker(st.session_state.template_bytes, "{{COLONNINE}}"),
            "{{FOTO_GALLERY}}": template_has_marker(st.session_state.template_bytes, "{{FOTO_GALLERY}}"),
            "{{DIAGRAMMA_IMPIANTO}}": template_has_marker(st.session_state.template_bytes, "{{DIAGRAMMA_IMPIANTO}}"),
            "{{ALLEGATI_SCHEDA_TECNICA}}": template_has_marker(st.session_state.template_bytes, "{{ALLEGATI_SCHEDA_TECNICA}}"),
        }

        missing_required = []

        # Layout: richiedi almeno una riga in incluso o escluso se la sezione è presente
        if present["{{LAYOUT_DESCRITTIVO}}"]:
            if not (layout_incluso.strip() or layout_escluso.strip()):
                missing_required.append("LAYOUT D'IMPIANTO (compila almeno Incluso o Escluso)")

        # Colonnine: richiedi almeno una colonnina se la sezione è presente
        if present["{{COLONNINE}}"]:
            if len(st.session_state.colonnine) == 0:
                missing_required.append("COLONNINE (aggiungi almeno una colonnina)")

        # Foto: richiedi almeno una foto se la sezione è presente
        if present["{{FOTO_GALLERY}}"]:
            if len(st.session_state.photos) == 0:
                missing_required.append("FOTO (carica almeno una foto)")

        # Diagramma: richiedi un diagramma se la sezione è presente
        if present["{{DIAGRAMMA_IMPIANTO}}"]:
            if not st.session_state.diagram_bytes:
                missing_required.append("DIAGRAMMA IMPIANTO (carica o disegna un diagramma)")

        # Allegati: richiedi almeno un allegato se la sezione è presente
        if present["{{ALLEGATI_SCHEDA_TECNICA}}"]:
            if len(st.session_state.allegati) == 0:
                missing_required.append("ALLEGATI SCHEDA TECNICA (carica almeno un PDF/DOCX)")

        # Ditta esecutrice: opzionale anche se il placeholder è presente
        # (se compili qualcosa, verrà inserita; altrimenti resterà vuota)

        if missing_required:
            st.error("Mancano dati obbligatori per queste sezioni del template:
- " + "
- ".join(missing_required))
            st.stop()
        progettista = ProgettistaData(
            nome=progettista_nome,
            indirizzo=progettista_indirizzo,
            cell=progettista_cell,
            email=progettista_email,
            piva=progettista_piva,
        )
        esecutrice = EsecutriceData(
            nome=esecutrice_nome,
            indirizzo=esecutrice_indirizzo,
            piva=esecutrice_piva,
        )

        data_obj = RelazioneData(
            luogo_data=luogo_data,
            committente_nome=committente_nome,
            sito_indirizzo=sito_indirizzo,
            sito_cap_citta=sito_cap_citta,
            oggetto=oggetto,
            distanza_m=int(distanza_m),
            potenza_impegnata_kw=float(potenza_impegnata_kw),
            cavo_tipo=cavo_tipo,
            cavo_lunghezza_m=int(cavo_lunghezza_m),
            ik_trifase_ka=float(ik_trifase_ka),
            ik_monofase_ka=float(ik_monofase_ka),
            layout_incluso=layout_incluso,
            layout_escluso=layout_escluso,
        )

        docx_bytes = generate_docx_bytes(
            template=st.session_state.template_bytes,
            data=data_obj,
            progettista=progettista,
            esecutrice=esecutrice,
            colonnine=st.session_state.colonnine,
            photos=st.session_state.photos,
            diagram_bytes=st.session_state.diagram_bytes,
            allegati=st.session_state.allegati,
        )
        st.session_state["last_docx_bytes"] = docx_bytes
        st.success("Documento generato. Ora puoi scaricarlo qui sotto.")

    if st.session_state.get("last_docx_bytes"):
        st.download_button(
            "Scarica DOCX",
            data=st.session_state["last_docx_bytes"],
            file_name=f"{filename_base}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
