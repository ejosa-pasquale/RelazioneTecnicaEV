import streamlit as st
import pandas as pd
import re
from datetime import date

from calcoli import corrente_da_potenza, caduta_tensione, verifica_tt_ra_idn, zs_massima_tn
from pdf_generator import genera_pdf_relazione_bytes

st.set_page_config(page_title="Relazione Tecnica DiCo – Impianto Elettrico", layout="wide")

st.title("Relazione Tecnica - Impianto Elettrico (supporto DiCo)")
st.caption("Interfaccia essenziale: inserisci i dati di progetto, compila tabelle (quadri/circuiti/EV/verifiche) e genera il PDF completo.")

PROGETTISTA_DEFAULT = """Ing. Pasquale Senese
Via Francesco Soave 30 - 20135 Milano (MI)
Cell: 340 5731381
Email: pasquale.senese@ingpec.eu
P.IVA: 14572980960
"""

CAVI_TIPO = ["FS17", "FG17", "FG16OR16", "FG16OM16"]

def _s(val: str) -> str:
    return (val or "").strip()

def _meaningful(val: str) -> bool:
    if val is None:
        return False
    s = str(val).strip()
    if not s:
        return False
    low = s.lower()
    bad = ["xxxx", "da inserire", "non compil", "n/a", "na", "—", "-"]
    return not any(b in low for b in bad)

st.sidebar.header("Parametri calcoli (globali)")
dv_lim = st.sidebar.number_input("Caduta di tensione max (%)", min_value=0.5, max_value=10.0, value=4.0, step=0.5)
cosphi_default = st.sidebar.number_input("cosφ default", min_value=0.50, max_value=1.00, value=0.95, step=0.01)
temp_amb = st.sidebar.number_input("Temperatura ambiente (°C)", min_value=-10, max_value=60, value=30, step=5)
st.sidebar.caption("Questi parametri influenzano le verifiche di sintesi in tabella circuiti.")

tab_doc, tab_impianto, tab_ev, tab_verifiche, tab_allegati = st.tabs(
    ["1) Dati documento", "2) Impianto & calcoli", "3) EV (se presente)", "4) Verifiche", "5) Allegati"]
)

with tab_doc:
    c1, c2, c3 = st.columns(3)
    with c1:
        committente_nome = st.text_input("Committente", value="")
        luogo = st.text_input("Luogo / Comune", value="")
    with c2:
        impianto_indirizzo = st.text_input("Luogo di installazione (indirizzo)", value="")
        oggetto_intervento = st.text_input("Oggetto intervento", value="")
    with c3:
        cod_progetto = st.text_input("Cod. progetto", value="")
        num_documento = st.text_input("N. documento", value="")
        revisione = st.text_input("Revisione", value="00")
        data_doc = st.date_input("Data documento", value=date.today())

    st.subheader("Dati tecnico-progettista / redattore")
    progettista_blocco = st.text_area("Dati progettista (come appariranno in PDF)", value=PROGETTISTA_DEFAULT, height=120)

    st.subheader("Dati di ingresso (rilievo / forniture)")
    c1, c2, c3 = st.columns(3)
    with c1:
        fonte_dati = st.text_input("Fonte dati (committente/impresa/gestore)", value="")
        data_conferma = st.date_input("Data conferma dati (se nota)", value=date.today())
    with c2:
        pod = st.text_input("POD (se disponibile)", value="")
        contatore_ubicazione = st.text_input("Contatore / misura ubicato in", value="")
    with c3:
        prescrizioni_enti = st.text_input("Prescrizioni Enti/Autorità (se presenti)", value="")

    st.subheader("Impostazioni impianto (generali)")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        tipologia_impianto = st.selectbox("Tipologia impianto", ["Manutenzione straordinaria", "Nuovo impianto", "Adeguamento", "Altro"])
        tipologia_altro = ""
        if tipologia_impianto == "Altro":
            tipologia_altro = st.text_input("Specificare tipologia", value="")
    with c2:
        sistema = st.selectbox("Sistema di distribuzione", ["TT", "TN", "Altro"])
        sistema_altro = ""
        if sistema == "Altro":
            sistema_altro = st.text_input("Specificare sistema", value="")
    with c3:
        tensione_freq = st.text_input("Tensione / Frequenza", value="230/400 V - 50 Hz")
        alimentazione = st.selectbox("Alimentazione", ["Monofase 230 V", "Trifase 400 V"])
    with c4:
        potenza_disponibile = st.text_input("Potenza impegnata / disponibile", value="")
        ambiente = st.multiselect("Ambientazioni particolari (se presenti)", ["Ordinario", "Esterno", "Autorimessa", "Cantiere", "Altro"], default=["Ordinario"])
        ambiente_altro = ""
        if "Altro" in ambiente:
            ambiente_altro = st.text_input("Specificare 'Altro'", value="")

with tab_impianto:
    st.subheader("Confini dell'intervento (opzionale, consigliato)")
    confini = st.text_area("Descrivi cosa è compreso/escluso e le interfacce con parti preesistenti/terze", value="", height=90)

    st.subheader("Quadri elettrici (sintesi)")
    quadri_df = pd.DataFrame([
        {"Quadro": "QG", "Ubicazione": "", "IP": "", "Interruttore generale (tipo/In)": "", "Differenziale generale (tipo/Idn)": ""}
    ])
    quadri = st.data_editor(
        quadri_df,
        use_container_width=True,
        num_rows="dynamic",
        key="quadri_editor",
    )

    st.subheader("Circuiti / linee / protezioni (sintesi + verifiche)")
    st.caption("Compila almeno: Linea, Destinazione, Lunghezza, Cavo, Protezione, Differenziale. Gli altri campi migliorano la relazione.")
    circuiti_df = pd.DataFrame([
        {"Linea": "L1", "Destinazione": "", "Posa": "", "L (m)": 0.0, "Cavo (tipo/sezione)": "FG16OM16 3Gx2.5", "Protezione (MT/MTD)": "MT 16A curva C", "Differenziale (tipo/Idn)": "Tipo A 30mA", "P (kW)": 0.0, "cosφ": cosphi_default}
    ])
    circuiti = st.data_editor(
        circuiti_df,
        use_container_width=True,
        num_rows="dynamic",
        key="circuiti_editor",
    )

    # Calcoli sintesi
    st.markdown("##### Sintesi calcoli (supporto)")
    calc_rows = []
    for _, r in circuiti.iterrows():
        linea = _s(r.get("Linea"))
        Lm = float(r.get("L (m)", 0) or 0)
        pkw = float(r.get("P (kW)", 0) or 0)
        cosphi = float(r.get("cosφ", cosphi_default) or cosphi_default)
        ib = None
        dv = None
        esito_dv = ""
        if pkw > 0:
            # stima monofase/trifase da alimentazione globale
            trifase = "Trifase" in alimentazione
            ib = corrente_da_potenza(pkw * 1000.0, 400.0 if trifase else 230.0, cosphi, trifase=trifase)
        # caduta tensione: serve sezione (cerchiamo ultima cifra nel campo)
        cavo = _s(r.get("Cavo (tipo/sezione)"))
        msec = re.search(r'(\d+(?:\.\d+)?)\s*$', cavo.replace(",", "."))
        S = float(msec.group(1)) if msec else None
        if ib is not None and S is not None and Lm > 0:
            trifase = "Trifase" in alimentazione
            dv = caduta_tensione(ib, Lm, S, trifase=trifase, cosphi=cosphi)
            esito_dv = "OK" if dv <= dv_lim else "KO"
        calc_rows.append({"Linea": linea, "Ib (A)": None if ib is None else round(ib, 2), "ΔV%": None if dv is None else round(dv, 2), "Esito ΔV": esito_dv})
    st.dataframe(pd.DataFrame(calc_rows), use_container_width=True)

with tab_ev:
    st.caption("Compila solo se l'intervento include infrastruttura di ricarica.")
    ev_df = pd.DataFrame([
        {"Tipo": "Wallbox/Colonnina", "Marca/Modello": "", "P (kW)": 0.0, "Alim.": "Monofase", "Modo": "3", "Connettore": "Tipo 2", "IP/IK": "", "RCD/RDC": "Tipo A + RDC-DD (o Tipo B secondo manuale)", "Note": ""}
    ])
    evse = st.data_editor(ev_df, use_container_width=True, num_rows="dynamic", key="ev_editor")

with tab_verifiche:
    st.subheader("Verifiche, prove e collaudi (CEI 64-8 Parte 6 / CEI 64-14)")
    ver_df = pd.DataFrame([
        {"Prova": "Esame a vista", "Esito": "", "Strumento": "", "Note": ""},
        {"Prova": "Continuità PE ed equipotenziale", "Esito": "", "Strumento": "", "Note": ""},
        {"Prova": "Resistenza di isolamento", "Esito": "", "Strumento": "", "Note": ""},
        {"Prova": "Prova differenziali (Idn/tempo)", "Esito": "", "Strumento": "", "Note": ""},
        {"Prova": "TT: misura Ra e coordinamento con Idn (se TT)", "Esito": "", "Strumento": "", "Note": ""},
        {"Prova": "TN: misura Zs e verifica intervento (se TN)", "Esito": "", "Strumento": "", "Note": ""},
        {"Prova": "Altre prove (SPD, emergenza, comandi, ecc.)", "Esito": "", "Strumento": "", "Note": ""},
    ])
    verifiche_tabella = st.data_editor(ver_df, use_container_width=True, num_rows="dynamic", key="ver_editor")
    st.caption("Suggerimento: per 'Esito' usa valori tipo 'positivo', 'negativo', 'non previsto', 'da eseguire'.")

with tab_allegati:
    st.subheader("Checklist documentale (fine report)")
    checklist_df = pd.DataFrame([
        {"Documento/Elaborato": "Dichiarazione di Conformità (DiCo) DM 37/08", "Stato": "", "Note": ""},
        {"Documento/Elaborato": "Relazione tipologica e materiali impiegati (DM 37/08)", "Stato": "", "Note": ""},
        {"Documento/Elaborato": "Schema/planimetria impianto (unifilare/multifilare) (DM 37/08)", "Stato": "", "Note": ""},
        {"Documento/Elaborato": "Verbali prove e misure (CEI 64-8 Parte 6 / CEI 64-14)", "Stato": "", "Note": ""},
        {"Documento/Elaborato": "Denuncia impianto di terra / verifiche periodiche (DPR 462/01) (se applicabile)", "Stato": "", "Note": ""},
        {"Documento/Elaborato": "Schede tecniche / dichiarazioni CE componenti principali", "Stato": "", "Note": ""},
        {"Documento/Elaborato": "Report fotografico essenziale", "Stato": "", "Note": ""},
    ])
    checklist = st.data_editor(checklist_df, use_container_width=True, num_rows="dynamic", key="check_editor")
    foto_files = st.file_uploader("Documentazione fotografica (max 6 immagini)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

st.divider()
st.subheader("Generazione PDF (report completo)")
st.caption("Il PDF include sempre il contenuto tecnico completo; i campi non compilati non vengono stampati.")
colA, colB = st.columns([1, 2])
with colA:
    cover_style = st.selectbox("Stile copertina", ["Engineering"], index=0, help="Unica modalità: report completo.")
with colB:
    note_generali = st.text_area("Note generali (opzionale)", value="", height=80)

if st.button("Genera PDF", type="primary"):
    # serializza foto
    foto_bytes = []
    if foto_files:
        for f in foto_files[:6]:
            foto_bytes.append({"name": f.name, "bytes": f.read()})

    # payload (solo campi necessari; vuoti non verranno stampati)
    payload = {
        "cover_style": "engineering",
        "committente_nome": _s(committente_nome),
        "luogo": _s(luogo),
        "impianto_indirizzo": _s(impianto_indirizzo),
        "oggetto_intervento": _s(oggetto_intervento),
        "tipologia_impianto": tipologia_altro if tipologia_impianto == "Altro" else tipologia_impianto,
        "sistema_distribuzione": sistema_altro if sistema == "Altro" else sistema,
        "tensione_freq": _s(tensione_freq),
        "alimentazione": alimentazione,
        "potenza_disponibile": _s(potenza_disponibile),
        "cod_progetto": _s(cod_progetto),
        "num_documento": _s(num_documento),
        "revisione": _s(revisione),
        "data_doc": str(data_doc),
        "progettista_blocco": _s(progettista_blocco),

        "fonte_dati": _s(fonte_dati),
        "data_conferma": str(data_conferma) if data_conferma else "",
        "pod": _s(pod),
        "contatore_ubicazione": _s(contatore_ubicazione),
        "prescrizioni_enti": _s(prescrizioni_enti),

        "ambienti": [a for a in ambiente if a != "Altro"] + ([_s(ambiente_altro)] if _meaningful(ambiente_altro) else []),
        "confini": _s(confini),
        "note_generali": _s(note_generali),

        "dv_lim": float(dv_lim),
        "cosphi_default": float(cosphi_default),

        "quadri": quadri.to_dict(orient="records"),
        "circuiti": circuiti.to_dict(orient="records"),
        "evse": evse.to_dict(orient="records"),
        "verifiche_tabella": verifiche_tabella.to_dict(orient="records"),
        "checklist": checklist.to_dict(orient="records"),
        "foto": foto_bytes,
    }

    pdf_bytes = genera_pdf_relazione_bytes(payload)
    st.download_button("Scarica PDF", data=pdf_bytes, file_name="Relazione_Tecnica_DiCo.pdf", mime="application/pdf")
