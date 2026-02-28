\
import datetime as dt
import json
from pathlib import Path
from typing import List, Optional

import streamlit as st
from streamlit_drawable_canvas import st_canvas

from generator import PhotoItem, RelazioneData, generate_docx, maybe_convert_to_pdf


APP_DIR = Path(__file__).parent
TEMPLATE_PATH = APP_DIR / "templates" / "relazione_base.docx"
OUTPUT_DIR = APP_DIR / "output"
PROFILES_DIR = APP_DIR / "profiles"


st.set_page_config(page_title="Generatore Relazione Progetto Elettrico", layout="wide")

st.title("Generatore relazione tecnica (progetto elettrico)")
st.caption("Compila i campi, aggiungi foto/diagramma e genera la relazione in .docx (e PDF se LibreOffice è disponibile).")

# ---------- Helpers ----------
def load_profile(name: str) -> dict:
    p = PROFILES_DIR / f"{name}.json"
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return {}

def save_profile(name: str, data: dict) -> None:
    PROFILES_DIR.mkdir(parents=True, exist_ok=True)
    p = PROFILES_DIR / f"{name}.json"
    p.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


# ---------- Sidebar: profilo ----------
with st.sidebar:
    st.header("Profilo progetto (dati che non cambiano)")
    existing_profiles = sorted([p.stem for p in PROFILES_DIR.glob("*.json")]) or ["default"]
    profile_name = st.selectbox("Seleziona profilo", options=existing_profiles, index=0)
    profile = load_profile(profile_name) or load_profile("default")

    st.caption("Suggerimento: imposta una volta i dati fissi e salva il profilo. Poi ti resteranno precompilati.")

    st.divider()
    st.header("Dati anagrafici")
    luogo = st.text_input("Luogo (es. Busnago)", value=profile.get("luogo", "Busnago"))
    data = st.date_input("Data", value=dt.date.today())
    luogo_data = f"{luogo}, {data.strftime('%d/%m/%Y')}"

    committente_nome = st.text_input("Committente (Nome / Condomìnio)", value=profile.get("committente_nome", "Nome"))
    sito_indirizzo = st.text_input("Indirizzo sito", value=profile.get("sito_indirizzo", "via della SS. Annunziata 32/A"))
    sito_cap_citta = st.text_input("CAP e Città (sito)", value=profile.get("sito_cap_citta", "55100 Lucca"))

    st.header("Richiedente")
    richiedente_nome = st.text_input("Ragione sociale", value=profile.get("richiedente_nome", "Telebit Spa"))
    richiedente_indirizzo = st.text_input("Indirizzo (richiedente)", value=profile.get("richiedente_indirizzo", "Via S. Rocco, 65"))
    richiedente_cap_citta = st.text_input("CAP e Città (richiedente)", value=profile.get("richiedente_cap_citta", "20874 Busnago"))

    st.header("Oggetto")
    oggetto = st.text_input("Oggetto intervento", value=profile.get("oggetto", "nuovo impianto di ricarica per veicolo elettrico"))

    st.header("Dati tecnici principali")
    distanza_m = st.number_input("Distanza punto origine (m)", min_value=1, value=int(profile.get("distanza_m", 60)), step=1)
    potenza_impegnata_kw = st.number_input("Potenza impegnata (kW)", min_value=0.5, value=float(profile.get("potenza_impegnata_kw", 4.0)), step=0.5)
    potenza_wallbox_kw = st.number_input("Potenza wallbox (kW)", min_value=0.5, value=float(profile.get("potenza_wallbox_kw", 7.4)), step=0.1)

    cavo_lunghezza_m = st.number_input("Lunghezza cavo (m)", min_value=1, value=int(profile.get("cavo_lunghezza_m", 60)), step=1)
    cavo_tipo = st.text_input("Tipo cavo", value=profile.get("cavo_tipo", "FG16OM16 3G6 0.6/1kV"))
    modello_wallbox = st.text_input("Modello wallbox", value=profile.get("modello_wallbox", "Cupra Charger Connect monofase"))

    st.header("Corto circuito (punto fornitura)")
    ik_trifase_ka = st.number_input("Ik trifase presunta (kA)", min_value=0.1, value=float(profile.get("ik_trifase_ka", 10.0)), step=0.1)
    ik_monofase_ka = st.number_input("Ik monofase presunta (kA)", min_value=0.1, value=float(profile.get("ik_monofase_ka", 6.0)), step=0.1)

    st.divider()
    st.subheader("Salva profilo")
    new_profile_name = st.text_input("Nome profilo (per salvare)", value=profile_name)
    if st.button("Salva / aggiorna profilo", use_container_width=True):
        save_profile(new_profile_name, {
            "luogo": luogo,
            "committente_nome": committente_nome,
            "sito_indirizzo": sito_indirizzo,
            "sito_cap_citta": sito_cap_citta,
            "richiedente_nome": richiedente_nome,
            "richiedente_indirizzo": richiedente_indirizzo,
            "richiedente_cap_citta": richiedente_cap_citta,
            "oggetto": oggetto,
            "distanza_m": distanza_m,
            "potenza_impegnata_kw": potenza_impegnata_kw,
            "potenza_wallbox_kw": potenza_wallbox_kw,
            "cavo_lunghezza_m": cavo_lunghezza_m,
            "cavo_tipo": cavo_tipo,
            "modello_wallbox": modello_wallbox,
            "ik_trifase_ka": ik_trifase_ka,
            "ik_monofase_ka": ik_monofase_ka,
        })
        st.success("Profilo salvato.")

    st.divider()
    st.subheader("Template")
    st.write("Di default uso il template basato sul tuo DOCX.")
    uploaded_template = st.file_uploader("Sostituisci template (.docx)", type=["docx"])
    if uploaded_template:
        TEMPLATE_PATH.parent.mkdir(parents=True, exist_ok=True)
        TEMPLATE_PATH.write_bytes(uploaded_template.getvalue())
        st.success("Template aggiornato.")
    st.caption("Per foto e diagramma inserisci nel template i placeholder: {{FOTO_GALLERY}} e {{DIAGRAMMA_IMPIANTO}}")


# ---------- Main content ----------
tab1, tab2, tab3 = st.tabs(["Dati & Generazione", "Foto di cantiere", "Diagramma impianto"])

# Foto (stato condiviso)
if "photos" not in st.session_state:
    st.session_state.photos = []  # List[PhotoItem]
if "diagram_bytes" not in st.session_state:
    st.session_state.diagram_bytes = None  # bytes

with tab2:
    st.subheader("Foto")
    st.write("Carica più immagini (JPG/PNG). Puoi aggiungere una didascalia per ciascuna.")
    uploads = st.file_uploader("Carica foto", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
    if uploads:
        # merge senza duplicati per filename
        existing = {p.filename for p in st.session_state.photos}
        for u in uploads:
            if u.name not in existing:
                st.session_state.photos.append(PhotoItem(filename=u.name, content=u.getvalue(), caption=""))
        st.success(f"Aggiunte {len(uploads)} foto (se non già presenti).")

    if st.session_state.photos:
        st.divider()
        st.write("Gestione foto (anteprima + didascalia):")
        for i, item in enumerate(list(st.session_state.photos)):
            c1, c2, c3 = st.columns([1, 2, 0.5])
            with c1:
                st.image(item.content, use_container_width=True)
            with c2:
                cap = st.text_input(f"Didascalia — {item.filename}", value=item.caption, key=f"cap_{i}")
                st.session_state.photos[i].caption = cap
            with c3:
                if st.button("Rimuovi", key=f"rm_{i}"):
                    st.session_state.photos.pop(i)
                    st.rerun()

        st.caption("Nel DOCX le foto verranno inserite dove trovi {{FOTO_GALLERY}} (griglia 2 colonne).")
    else:
        st.info("Nessuna foto caricata (opzionale).")


with tab3:
    st.subheader("Diagramma impianto")
    st.write("Puoi **caricare** un diagramma (PNG/JPG) oppure **disegnarlo** rapidamente qui.")
    mode = st.radio("Metodo", ["Carica immagine", "Disegna"], horizontal=True)

    if mode == "Carica immagine":
        diag_up = st.file_uploader("Carica diagramma", type=["jpg", "jpeg", "png"], accept_multiple_files=False)
        if diag_up:
            st.session_state.diagram_bytes = diag_up.getvalue()
            st.image(st.session_state.diagram_bytes, caption="Diagramma selezionato", use_container_width=True)

    else:
        st.caption("Disegno semplice: usa mouse/trackpad per tracciare linee e note.")
        canvas = st_canvas(
            fill_color="rgba(255, 255, 255, 0)",
            stroke_width=3,
            stroke_color="#000000",
            background_color="#FFFFFF",
            height=450,
            drawing_mode="freedraw",
            key="canvas",
        )
        colA, colB = st.columns([1, 1])
        with colA:
            if st.button("Usa disegno come diagramma", type="primary"):
                if canvas.image_data is not None:
                    # convert RGBA numpy -> PNG bytes
                    import PIL.Image
                    import numpy as np
                    img = PIL.Image.fromarray(canvas.image_data.astype("uint8"), mode="RGBA").convert("RGB")
                    import io
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    st.session_state.diagram_bytes = buf.getvalue()
                    st.success("Diagramma salvato.")
        with colB:
            if st.session_state.diagram_bytes:
                st.image(st.session_state.diagram_bytes, caption="Diagramma pronto per l'inserimento", use_container_width=True)

    st.caption("Nel DOCX il diagramma verrà inserito dove trovi {{DIAGRAMMA_IMPIANTO}}.")


with tab1:
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("Anteprima dati")
        st.write(
            {
                "luogo_data": luogo_data,
                "committente_nome": committente_nome,
                "sito_indirizzo": sito_indirizzo,
                "sito_cap_citta": sito_cap_citta,
                "richiedente_nome": richiedente_nome,
                "oggetto": oggetto,
                "distanza_m": distanza_m,
                "potenza_impegnata_kw": potenza_impegnata_kw,
                "potenza_wallbox_kw": potenza_wallbox_kw,
                "cavo": f"{cavo_lunghezza_m} m - {cavo_tipo}",
                "wallbox": modello_wallbox,
                "ik_trifase_ka": ik_trifase_ka,
                "ik_monofase_ka": ik_monofase_ka,
                "foto_caricate": len(st.session_state.photos),
                "diagramma": "Sì" if st.session_state.diagram_bytes else "No",
            }
        )

    with col2:
        st.subheader("Generazione documento")
        st.warning(
            "Per inserire **foto** e **diagramma** nel file Word, apri il template e inserisci i placeholder:\n"
            "- {{FOTO_GALLERY}}\n"
            "- {{DIAGRAMMA_IMPIANTO}}\n\n"
            "Esattamente così, nella posizione desiderata."
        )

        filename_base = st.text_input("Nome file", value="relazione_progetto_elettrico")
        generate = st.button("Genera relazione", type="primary", use_container_width=True)

        if generate:
            data_obj = RelazioneData(
                luogo_data=luogo_data,
                committente_nome=committente_nome,
                sito_indirizzo=sito_indirizzo,
                sito_cap_citta=sito_cap_citta,
                richiedente_nome=richiedente_nome,
                richiedente_indirizzo=richiedente_indirizzo,
                richiedente_cap_citta=richiedente_cap_citta,
                oggetto=oggetto,
                distanza_m=int(distanza_m),
                potenza_impegnata_kw=float(potenza_impegnata_kw),
                potenza_wallbox_kw=float(potenza_wallbox_kw),
                cavo_tipo=cavo_tipo,
                cavo_lunghezza_m=int(cavo_lunghezza_m),
                modello_wallbox=modello_wallbox,
                ik_trifase_ka=float(ik_trifase_ka),
                ik_monofase_ka=float(ik_monofase_ka),
            )

            out_docx = OUTPUT_DIR / f"{filename_base}.docx"
            generate_docx(
                TEMPLATE_PATH,
                out_docx,
                data_obj,
                photos=st.session_state.photos,
                diagram_bytes=st.session_state.diagram_bytes,
            )

            st.success("Relazione generata.")
            st.download_button(
                "Scarica DOCX",
                data=out_docx.read_bytes(),
                file_name=out_docx.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

            pdf_path = OUTPUT_DIR / f"{filename_base}.pdf"
            pdf_ok = maybe_convert_to_pdf(out_docx, pdf_path)
            if pdf_ok:
                st.download_button(
                    "Scarica PDF",
                    data=pdf_ok.read_bytes(),
                    file_name=pdf_ok.name,
                    mime="application/pdf",
                    use_container_width=True,
                )
            else:
                st.caption("PDF non disponibile (LibreOffice non trovato).")
