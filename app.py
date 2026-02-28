\
import datetime as dt
from pathlib import Path

import streamlit as st

from generator import RelazioneData, generate_docx, maybe_convert_to_pdf


APP_DIR = Path(__file__).parent
TEMPLATE_PATH = APP_DIR / "templates" / "relazione_base.docx"
OUTPUT_DIR = APP_DIR / "output"


st.set_page_config(page_title="Generatore Relazione Progetto Elettrico", layout="wide")

st.title("Generatore relazione tecnica (progetto elettrico)")
st.caption("Compila i campi e genera la relazione in .docx (e PDF se LibreOffice è disponibile).")

with st.sidebar:
    st.header("Dati anagrafici")
    luogo = st.text_input("Luogo (es. Busnago)", value="Busnago")
    data = st.date_input("Data", value=dt.date.today())
    luogo_data = f"{luogo}, {data.strftime('%d/%m/%Y')}"

    committente_nome = st.text_input("Committente (Nome / Condomìnio)", value="Nome")
    sito_indirizzo = st.text_input("Indirizzo sito", value="via della SS. Annunziata 32/A")
    sito_cap_citta = st.text_input("CAP e Città (sito)", value="55100 Lucca")

    st.header("Richiedente")
    richiedente_nome = st.text_input("Ragione sociale", value="Telebit Spa")
    richiedente_indirizzo = st.text_input("Indirizzo (richiedente)", value="Via S. Rocco, 65")
    richiedente_cap_citta = st.text_input("CAP e Città (richiedente)", value="20874 Busnago")

    st.header("Oggetto")
    oggetto = st.text_input("Oggetto intervento", value="nuovo impianto di ricarica per veicolo elettrico")

    st.header("Dati tecnici principali")
    distanza_m = st.number_input("Distanza punto origine (m)", min_value=1, value=60, step=1)
    potenza_impegnata_kw = st.number_input("Potenza impegnata (kW)", min_value=0.5, value=4.0, step=0.5)
    potenza_wallbox_kw = st.number_input("Potenza wallbox (kW)", min_value=0.5, value=7.4, step=0.1)

    cavo_lunghezza_m = st.number_input("Lunghezza cavo (m)", min_value=1, value=60, step=1)
    cavo_tipo = st.text_input("Tipo cavo", value="FG16OM16 3G6 0.6/1kV")
    modello_wallbox = st.text_input("Modello wallbox", value="Cupra Charger Connect monofase")

    st.header("Corto circuito (punto fornitura)")
    ik_trifase_ka = st.number_input("Ik trifase presunta (kA)", min_value=0.1, value=10.0, step=0.1)
    ik_monofase_ka = st.number_input("Ik monofase presunta (kA)", min_value=0.1, value=6.0, step=0.1)

    st.divider()
    st.subheader("Template")
    st.write("Di default uso il template basato sul tuo DOCX.")
    uploaded_template = st.file_uploader("Sostituisci template (.docx)", type=["docx"])
    if uploaded_template:
        TEMPLATE_PATH.parent.mkdir(parents=True, exist_ok=True)
        TEMPLATE_PATH.write_bytes(uploaded_template.getvalue())
        st.success("Template aggiornato.")


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
        }
    )

with col2:
    st.subheader("Generazione documento")
    st.info("Nota: alcune parti in **textbox/forme Word** potrebbero non essere sostituite automaticamente da python-docx.")
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
        generate_docx(TEMPLATE_PATH, out_docx, data_obj)

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
