\
import datetime as dt
import json
from pathlib import Path

import streamlit as st
from streamlit_drawable_canvas import st_canvas

from generator import PhotoItem, RelazioneData, generate_docx_bytes


APP_DIR = Path(__file__).parent
DEFAULT_TEMPLATE_PATH = APP_DIR / "templates" / "relazione_base.docx"
PROFILES_DIR = APP_DIR / "profiles"

st.set_page_config(page_title="Generatore Relazione Progetto Elettrico", layout="wide")
st.title("Generatore relazione tecnica (progetto elettrico)")
st.caption("Compila i campi, aggiungi foto/diagramma e scarica il DOCX generato (download sempre disponibile).")


# ---------- Helpers ----------
def load_profile(name: str) -> dict:
    p = PROFILES_DIR / f"{name}.json"
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return {}

def save_profile(name: str, data: dict) -> None:
    # Nota: su Streamlit Cloud la repo può essere read-only. Proviamo; se fallisce, avvisiamo.
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


# ---------- Sidebar ----------
with st.sidebar:
    st.header("Profilo progetto (dati fissi)")
    existing_profiles = sorted([p.stem for p in PROFILES_DIR.glob("*.json")]) or ["default"]
    profile_name = st.selectbox("Seleziona profilo", options=existing_profiles, index=0)
    profile = load_profile(profile_name) or load_profile("default")

    st.divider()
    st.header("Dati anagrafici")
    luogo = st.text_input("Luogo (es. Busnago)", value=profile.get("luogo", "Busnago"))
    data = st.date_input("Data", value=dt.date.today())
    luogo_data = f"{luogo}, {data.strftime('%d/%m/%Y')}"

    committente_nome = st.text_input("Committente (Nome / Condomìnio)", value=profile.get("committente_nome", "Nome"))
    sito_indirizzo = st.text_input("Indirizzo sito", value=profile.get("sito_indirizzo", "via della SS. Annunziata 32/A"))
    sito_cap_citta = st.text_input("CAP e Città (sito)", value=profile.get("sito_cap_citta", "55100 Lucca"))

    st.header("Richiedente")
    richiedente_nome = st.text_input("Richiedente", value=profile.get("richiedente_nome", "Ing Pasquale Senese"))
    richiedente_indirizzo = st.text_input("Indirizzo (opzionale)", value=profile.get("richiedente_indirizzo", ""))
    richiedente_cap_citta = st.text_input("CAP e Città (opzionale)", value=profile.get("richiedente_cap_citta", ""))

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
        try:
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
        except Exception as e:
            st.warning("Non riesco a scrivere su disco (tipico su Streamlit Cloud). Il profilo non è stato salvato.")
            st.caption(str(e))

    st.divider()
    st.subheader("Template (senza logo Telebit)")
    st.write("Uso il template **no-logo** come base. Puoi comunque caricarne uno alternativo.")
    uploaded_template = st.file_uploader("Sostituisci template (.docx)", type=["docx"])
    if uploaded_template:
        st.session_state.template_bytes = uploaded_template.getvalue()
        st.success("Template caricato (in memoria).")
    st.caption("Per foto e diagramma inserisci nel template: {{FOTO_GALLERY}} e {{DIAGRAMMA_IMPIANTO}}")


# ---------- Tabs ----------
tab1, tab2, tab3 = st.tabs(["Dati & Download", "Foto di cantiere", "Diagramma impianto"])

with tab2:
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
        st.caption("Nel DOCX le foto entrano dove trovi {{FOTO_GALLERY}}.")
    else:
        st.info("Nessuna foto caricata (opzionale).")


with tab3:
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
                import numpy as np
                import io
                img = PIL.Image.fromarray(canvas.image_data.astype("uint8"), mode="RGBA").convert("RGB")
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                st.session_state.diagram_bytes = buf.getvalue()
                st.success("Diagramma salvato.")
        if st.session_state.diagram_bytes:
            st.image(st.session_state.diagram_bytes, caption="Diagramma pronto", use_container_width=True)

    st.caption("Nel DOCX il diagramma entra dove trovi {{DIAGRAMMA_IMPIANTO}}.")


with tab1:
    st.subheader("Generazione")
    st.warning(
        "Per inserire **foto** e **diagramma** nel Word, nel template devono esserci i placeholder:\n"
        "- {{FOTO_GALLERY}}\n"
        "- {{DIAGRAMMA_IMPIANTO}}"
    )

    filename_base = st.text_input("Nome file", value="relazione_progetto_elettrico")
    if st.button("Genera e prepara download", type="primary", use_container_width=True):
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

        docx_bytes = generate_docx_bytes(
            template=st.session_state.template_bytes,
            data=data_obj,
            photos=st.session_state.photos,
            diagram_bytes=st.session_state.diagram_bytes,
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
