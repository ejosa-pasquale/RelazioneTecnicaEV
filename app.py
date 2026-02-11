import streamlit as st
import pandas as pd
from datetime import date

from calcoli import corrente_da_potenza, caduta_tensione, verifica_tt_ra_idn, zs_massima_tn
from pdf_generator import genera_pdf_relazione_bytes

st.set_page_config(page_title="Relazione Tecnica DiCo – Impianti Elettrici", layout="wide")

st.title("Relazione Tecnica - Impianto Elettrico (Allegato alla DiCo)")
st.caption("Compilazione guidata (stile v7) + calcoli essenziali + generazione PDF.")

PROGETTISTA_BLOCCO = """Ing. Pasquale Senese
Via Francesco Soave 30 - 20151 Milano
Cell: 340 5731381
Email: pasquale.senese@ingpec.eu
P.IVA: 14572980960
"""

CAVI_TIPO = ["FS17", "FG17", "FG16OR16", "FG16OM16"]

with st.sidebar:
    st.header("Parametri calcoli (sintesi)")
    dv_lim = st.number_input("Caduta di tensione max (%)", min_value=1.0, max_value=10.0, value=4.0, step=0.5)
    ul_tt = st.number_input("UL sistema TT (V) – criterio Ra·Idn ≤ UL", min_value=25.0, max_value=100.0, value=50.0, step=5.0)
    st.divider()
    st.markdown("**Nota**: calcoli di sintesi (supporto alla DiCo). Per casi complessi usare calcolo/progetto dedicato.")

# =========================
# DATI IDENTIFICATIVI
# =========================
st.subheader("Dati identificativi documento")

c1, c2, c3 = st.columns(3)
with c1:
    committente = st.text_input("Committente", "XXXX (Inserire)")
    luogo = st.text_input("Luogo di installazione (indirizzo completo)", "XXXX (Inserire indirizzo completo)")
    oggetto = st.text_input("Oggetto intervento (descrizione sintetica)", "XXXX (Inserire descrizione sintetica dell’intervento)")
with c2:
    tipologia = st.selectbox("Tipologia impianto", ["Nuova realizzazione", "Ampliamento", "Trasformazione", "Manutenzione straordinaria"], index=3)
    sistema = st.selectbox("Sistema di distribuzione", ["TT", "TN-S", "TN-C-S", "IT"], index=0)
    alimentazione = st.selectbox("Alimentazione", ["Monofase 230 V", "Trifase 400 V"], index=1)
with c3:
    tensione = st.text_input("Tensione/Frequenza", "230/400 V - 50 Hz")
    potenza_disp_kw = st.text_input("Potenza impegnata / disponibile", "XXXX (Inserire)")
    n_doc = st.text_input("N. documento", "XXXX (Inserire)")
    revisione = st.text_input("Revisione", "00")
    data_doc = st.date_input("Data", value=date.today())

st.divider()

# =========================
# SOGGETTI COINVOLTI
# =========================
st.subheader("Soggetti coinvolti")

c1, c2 = st.columns(2)
with c1:
    st.markdown("**Impresa installatrice**")
    impresa = st.text_input("Ragione sociale", "XXXX (Inserire)")
    impresa_sede = st.text_input("Sede legale", "XXXX (Inserire)")
    impresa_piva = st.text_input("P.IVA / C.F.", "XXXX (Inserire)")
    impresa_rea = st.text_input("N. iscrizione CCIAA / REA", "XXXX (Inserire)")
    impresa_resp = st.text_input("Responsabile tecnico", "XXXX (Inserire)")
    impresa_cont = st.text_input("Recapiti", "XXXX (Inserire)")
with c2:
    st.markdown("**Progettista / Tecnico redattore (se applicabile)**")
    st.code(PROGETTISTA_BLOCCO)

st.divider()

# =========================
# DATI TECNICI MINIMI
# =========================
st.subheader("Dati tecnici minimi (da compilare)")

c1, c2, c3, c4 = st.columns(4)
with c1:
    pod = st.text_input("POD / punto di consegna", "XXXX (Inserire)")
with c2:
    contatore_ubi = st.text_input("Contatore ubicato in", "XXXX (Inserire)")
with c3:
    potenza_prev_kw = st.number_input("Potenza prevista/servita (kW) – per stima Ib", min_value=0.5, max_value=500.0, value=6.0, step=0.5)
with c4:
    cosphi = st.number_input("cosφ (se noto)", min_value=0.3, max_value=1.0, value=0.95, step=0.01)

Ib = corrente_da_potenza(potenza_prev_kw, alimentazione, cosphi=cosphi)
st.info(f"Corrente di impiego indicativa Ib ≈ **{Ib:.1f} A** (stima da potenza {potenza_prev_kw:.1f} kW, cosφ={cosphi:.2f}).")

ambienti = st.multiselect(
    "Destinazione d’uso / ambienti (checklist)",
    ["Ordinario", "Bagno", "Esterno", "Locale tecnico", "Autorimessa", "Maggior rischio incendio", "Cantiere", "Altro"],
    default=["Ordinario"],
)
amb_altro = ""
if "Altro" in ambienti:
    amb_altro = st.text_input("Specificare 'Altro'", "XXXX (Inserire)")

st.divider()

# =========================
# CONFINE INTERVENTO
# =========================
st.subheader("Confini dell’intervento e interfacce")

c1, c2 = st.columns(2)
with c1:
    compresi = st.text_area("L’intervento comprende", "XXXX (Inserire elenco sintetico delle opere incluse).", height=120)
with c2:
    esclusi = st.text_area("Sono esclusi", "XXXX (Inserire, es. parti preesistenti non modificate, linee a monte, apparecchiature non comprese).", height=120)

integrazione = st.selectbox("Integrazione con impianto esistente", ["Sì", "No"], index=0)
integrazione_note = ""
if integrazione == "Sì":
    integrazione_note = st.text_area("Descrizione e condizioni riscontrate/limiti di intervento", "XXXX (Inserire).", height=90)

st.divider()

# =========================
# QUADRI
# =========================
st.subheader("Quadri elettrici e distribuzione (tabella sintetica)")

default_quadri = pd.DataFrame([
    {"Quadro":"QG", "Ubicazione":"XXXX (Inserire)", "IP":"XX", "Interruttore generale (tipo/In)":"XXXX (Inserire)", "Differenziale generale (tipo/Idn, se presente)":"XXXX (Inserire)"},
])
quadri_df = st.data_editor(default_quadri, num_rows="dynamic", use_container_width=True, key="quadri")

st.divider()

# =========================
# LINEE / CIRCUITI
# =========================
st.subheader("Circuiti, cavi e protezioni (con calcolo ΔV)")

st.caption("Per ciascun circuito: scegli **Tipo cavo (FS17/FG17/FG16OR16/FG16OM16)**, sezione, protezione e differenziale. "
           "Il calcolo automatico mostra ΔV% e un esito sintetico.")

default_linee = pd.DataFrame([
    {"Circuito/Linea":"L1", "Destinazione/Utilizzo":"Prese", "Potenza_kW":2.0, "Posa":"XXXX", "Lunghezza_m":25,
     "Tipo_cavo":"FG16OM16", "Formazione":"3G", "Sezione_mm2":2.5,
     "Protezione (MT/MTD)":"MT 16A curva C", "Curva":"C", "In_A":16,
     "Differenziale (tipo/Idn)":"Tipo A 30mA", "Idn_A":0.03,
     "Ra_Ohm (solo TT)":30.0},
])

linee_df = st.data_editor(
    default_linee,
    num_rows="dynamic",
    use_container_width=True,
    key="linee",
    column_config={
        "Tipo_cavo": st.column_config.SelectboxColumn("Tipo cavo", options=CAVI_TIPO, required=True),
        "Formazione": st.column_config.TextColumn("Formazione (es. 3G / 5G)", help="Esempio: 3G per monofase+PE, 5G per trifase+N+PE."),
    }
)

def valuta_linea(row):
    p = float(row.get("Potenza_kW") or 0.0)
    l = float(row.get("Lunghezza_m") or 0.0)
    s = float(row.get("Sezione_mm2") or 0.0)
    curva = str(row.get("Curva") or "C")
    InA = float(row.get("In_A") or 0.0)

    ib_linea = corrente_da_potenza(p, alimentazione, cosphi=cosphi) if p > 0 else Ib

    dv = caduta_tensione(ib_linea, l, s, alimentazione, cosphi=cosphi)
    esito = "OK" if dv.delta_v_percent <= dv_lim else "ΔV"

    note = []
    if sistema == "TT":
        ra = float(row.get("Ra_Ohm (solo TT)") or 0.0)
        idn = float(row.get("Idn_A") or 0.0)
        if ra > 0 and idn > 0:
            ok_tt = verifica_tt_ra_idn(ra, idn, ul=ul_tt)
            note.append("TT OK" if ok_tt else "TT NO")
            if not ok_tt:
                esito = "TT"
    else:
        if InA > 0:
            zs_max = zs_massima_tn(230.0, curva, InA)
            note.append(f"Zs_max≈{zs_max:.2f}Ω")

    return dv.delta_v_percent, esito, "; ".join(note)

out = [valuta_linea(r) for _, r in linee_df.iterrows()]

linee_df_calc = linee_df.copy()
linee_df_calc["ΔV_%"] = [round(x[0], 2) for x in out]
linee_df_calc["Esito"] = [x[1] for x in out]
linee_df_calc["Note"] = [x[2] for x in out]

st.dataframe(linee_df_calc, use_container_width=True)

st.divider()

# =========================
# SICUREZZA / TERRA / SPD
# =========================
st.subheader("Sicurezza elettrica, terra, SPD (sintesi)")

c1, c2 = st.columns(2)
with c1:
    terra_cfg = st.selectbox("Configurazione impianto di terra", ["Nuovo", "Esistente verificato", "Esistente non oggetto di intervento (da motivare)"], index=1)
    dispersore = st.text_input("Dispersore (descrizione)", "XXXX (Inserire)")
    equipot = st.selectbox("Collegamenti equipotenziali principali", ["Presenti", "Parziali", "Assenti (da adeguare/indicare)"], index=0)
with c2:
    spd_esito = st.selectbox("Protezione contro sovratensioni (SPD) – esito", ["Non previsto", "Previsto", "Presente preesistente"], index=0)
    spd_tipo = st.multiselect("Se installato: tipologia SPD", ["Tipo 1", "Tipo 2", "Tipo 3"], default=[])
    spd_quadro = st.text_input("Quadro di installazione SPD (se pertinente)", "XXXX (Inserire)")
    spd_caratt = st.text_input("Caratteristiche principali SPD (se pertinente)", "XXXX (Inserire)")

st.divider()

# =========================
# VERIFICHE
# =========================
st.subheader("Verifiche, prove e collaudi (registro sintetico)")

ver_df = pd.DataFrame([
    {"Prova / Verifica":"Esame a vista", "Esito":"XXXX (Inserire)", "Strumento":"XXXX (Inserire)", "Note":"XXXX (Inserire)"},
    {"Prova / Verifica":"Continuità PE ed equipotenziale", "Esito":"XXXX (Inserire)", "Strumento":"XXXX (Inserire)", "Note":"XXXX (Inserire)"},
    {"Prova / Verifica":"Resistenza di isolamento", "Esito":"XXXX (Inserire)", "Strumento":"XXXX (Inserire)", "Note":"XXXX (Inserire)"},
    {"Prova / Verifica":"Prova differenziali (Idn/tempo)", "Esito":"XXXX (Inserire)", "Strumento":"XXXX (Inserire)", "Note":"XXXX (Inserire)"},
    {"Prova / Verifica":"Polarità / sequenza fasi (se pertinente)", "Esito":"XXXX (Inserire)", "Strumento":"XXXX (Inserire)", "Note":"XXXX (Inserire)"},
    {"Prova / Verifica":"TT: misura Ra e coordinamento con Idn (se TT)", "Esito":"XXXX (Inserire)", "Strumento":"XXXX (Inserire)", "Note":"XXXX (Inserire)"},
    {"Prova / Verifica":"TN: misura Zs e verifica intervento (se TN)", "Esito":"XXXX (Inserire)", "Strumento":"XXXX (Inserire)", "Note":"XXXX (Inserire)"},
    {"Prova / Verifica":"Altre prove (SPD, emergenza, comandi, ecc.)", "Esito":"XXXX (Inserire)", "Strumento":"XXXX (Inserire)", "Note":"XXXX (Inserire)"},
])
ver_df = st.data_editor(ver_df, num_rows="dynamic", use_container_width=True, key="verifiche")

st.divider()

# =========================
# GENERAZIONE PDF
# =========================
st.subheader("Genera PDF")

premessa = f"""La presente relazione tecnica è redatta a supporto della Dichiarazione di Conformità (DiCo) dell’impianto elettrico realizzato presso: {luogo}, per conto della committenza: {committente}. Il documento descrive le scelte progettuali e realizzative, i criteri di dimensionamento e le verifiche previste/effettuate, in conformità alle normative vigenti e alle regole dell’arte.

Nota sul ruolo: la presente Relazione Tecnica è redatta a supporto della Dichiarazione di Conformità (DiCo) rilasciata dall’Impresa installatrice ai sensi del D.M. 37/2008. La Relazione descrive l’impianto e le verifiche previste/effettuate e non sostituisce la DiCo né i relativi allegati obbligatori. Il Progettista/Tecnico redattore assume responsabilità nei limiti dell’incarico conferito (progettazione e/o verifica/collaudo, ove formalmente previsto).

L’intervento riguarda: {oggetto}

Assunzioni e fonti: le informazioni relative a fornitura elettrica (POD, potenza disponibile/contrattuale, caratteristiche del punto di consegna), destinazione d’uso e condizioni di esercizio sono state fornite da XXXX (Committente/Impresa/Gestore) e/o rilevate in sito in data {data_doc.strftime('%d/%m/%Y')}. Eventuali parti preesistenti non oggetto di intervento sono indicate nel paragrafo “Confini dell’intervento”.
"""

norme = """Si riportano i principali riferimenti legislativi e normativi applicabili (elenco non esaustivo):

• D.M. 22/01/2008 n. 37 (Regolamento per l’installazione degli impianti all’interno degli edifici).
• Legge 01/03/1968 n. 186 (regola dell’arte e norme CEI).
• D.Lgs. 09/04/2008 n. 81 e s.m.i. (sicurezza nei luoghi di lavoro).
• D.P.R. 22/10/2001 n. 462 (messa a terra e protezione scariche atmosferiche, ove applicabile).
• Norme CEI applicabili, con particolare riferimento a: CEI 64-8, CEI 0-2, CEI 0-21/0-16 (se applicabili), CEI EN 61439 (quadri), CEI EN 60529 (IP), CEI 64-14 (verifiche), CEI 0-10 (manutenzione), CEI 81-10 (fulmini/LPS, se applicabile).
• Regolamento Prodotti da Costruzione (UE) 305/2011 (CPR) e norme CEI-UNEL per i cavi (ove applicabile).

Eventuali ulteriori prescrizioni di Enti/Autorità locali (VV.F., gestore di rete, regolamenti condominiali, ecc.): XXXX (Inserire).
"""

amb_txt = ", ".join([a for a in ambienti if a != "Altro"])
if "Altro" in ambienti:
    amb_txt += f", Altro: {amb_altro}"

dati_tecnici = f"""Tipo sistema di distribuzione: {sistema}. Tensione nominale: {tensione}. Potenza disponibile/contrattuale: {potenza_disp_kw}.
POD: {pod} – contatore ubicato in: {contatore_ubi}.
Alimentazione: {alimentazione}. Potenza prevista/servita (stima): {potenza_prev_kw:.1f} kW (Ib indicativa ≈ {Ib:.1f} A a cosφ={cosphi:.2f}).
Ambientazioni particolari (se presenti): {amb_txt}.
"""

descrizione_impianto = f"""Il sito di intervento è ubicato in {luogo}. L’impianto è alimentato in bassa tensione dal punto di consegna del Distributore (POD: {pod}), tramite contatore/quadretto di misura ubicato in {contatore_ubi}.

Tipo sistema di distribuzione: {sistema}. Tensione nominale: {tensione}. Potenza disponibile/contrattuale: {potenza_disp_kw}.

La ripartizione e distribuzione interna avviene mediante linee in cavo conforme CEI/UNEL e componenti marcati CE (e, ove disponibile, IMQ o equivalente).

Scopo dell’intervento (descrizione sintetica): {oggetto}

Le opere impiantistiche previste comprendono, in funzione dell’intervento, la realizzazione e/o modifica di linee di alimentazione dedicate, installazione di prese e punti di utilizzo, posa di tubazioni/canalizzazioni, installazione o adeguamento di quadri elettrici (generale e/o di zona), apparecchi di protezione e comando, morsetterie e accessori, nonché collegamenti al sistema di protezione (PE) e ai collegamenti equipotenziali.

La distribuzione in bassa tensione avviene a partire dal punto di consegna/contatore e dai quadri esistenti, con realizzazione di linee in cavo e condutture posate in tubazioni/canalizzazioni idonee. Le linee di distribuzione e i circuiti terminali sono organizzati per destinazione d’uso (prese, illuminazione, ausiliari, ecc.), privilegiando la separazione funzionale e la facilità di manutenzione. I conduttori sono identificati secondo codifica colori (PE giallo-verde, N blu, fasi marrone/nero/grigio) e marcatura/etichettatura.
"""

confini_txt = f"""L’intervento comprende: {compresi}

Sono esclusi: {esclusi}

Integrazione con impianto esistente: {integrazione}. {("Descrizione e limiti: " + integrazione_note) if integrazione == "Sì" else ""}
"""

sicurezza = f"""La protezione contro i contatti diretti è assicurata tramite isolamento delle parti attive, involucri/barriere con grado di protezione adeguato e corretta posa delle condutture.

La protezione contro i contatti indiretti è assicurata mediante interruzione automatica dell’alimentazione, in accordo con CEI 64-8, tramite dispositivi differenziali e/o magnetotermici coordinati con l’impianto di terra (nei sistemi TT) o con il conduttore di protezione (nei sistemi TN).

Per i circuiti prese a uso generale è normalmente adottata protezione differenziale ad alta sensibilità: Idn = XXXX (tipicamente 30 mA), tipo: XXXX (AC / A / F / B in funzione dei carichi).

Configurazione impianto di terra: {terra_cfg}. Dispersore: {dispersore}. Collegamenti equipotenziali principali: {equipot}.

Protezione contro le sovratensioni (SPD) – esito: {spd_esito}. {("Tipologia: " + ", ".join(spd_tipo) + " – ") if spd_tipo else ""}quadro: {spd_quadro}. Caratteristiche: {spd_caratt}.

Caduta di tensione: verificata entro il limite adottato in progetto: {dv_lim:.1f}%.
"""

verifiche = """Ad ultimazione dei lavori, l’impianto è sottoposto alle verifiche previste dalla CEI 64-8 (Parte 6) e dalla CEI 64-14, con esecuzione e registrazione delle prove strumentali pertinenti al sistema di distribuzione (TT/TN) e alla tipologia di impianto. In particolare:

"""
for _, r in ver_df.iterrows():
    verifiche += f"• {r.get('Prova / Verifica','')}: {r.get('Esito','')} – Strumento: {r.get('Strumento','')} – Note: {r.get('Note','')}\n"

manutenzione = """L’esercizio e la manutenzione devono essere eseguiti da personale qualificato, secondo quanto indicato dai costruttori e dalla normativa tecnica (CEI 0-10). Si raccomanda l’esecuzione di controlli periodici e la registrazione degli interventi. Eventuali prescrizioni specifiche (es. per apparecchiature particolari): XXXX (Inserire)."""

allegati = """Completano la presente relazione e/o la DiCo i seguenti allegati. Gli allegati indicati come 'Obbligatorio' devono essere presenti nella documentazione finale; gli altri sono da allegare se disponibili o se pertinenti all’intervento.

- Schema unifilare / multifilare dei quadri interessati: Obbligatorio.
- Planimetrie con tracciati e posizionamento componenti: Obbligatorio.
- Elenco linee/circuiti con cavo e protezione (tabella circuiti): Obbligatorio.
- Verbali e report delle misure e prove strumentali: Obbligatorio.
- Schede tecniche principali componenti (quadri, interruttori, SPD, ecc.): Se disponibile.
- Dichiarazioni/Marcature CE (ed eventuale IMQ) dei materiali: Se disponibile.
- Documentazione impianto di terra (schemi, misure, foto morsetti): Se pertinente.
- Documentazione SPD / LPS (se presenti): Se pertinente.
- Report fotografico essenziale (quadri, targhette, collegamenti di terra, punti significativi): Consigliato.
"""

if st.button("Genera PDF"):
    quadri_list = []
    for _, q in quadri_df.iterrows():
        quadri_list.append({
            "Quadro": q.get("Quadro",""),
            "Ubicazione": q.get("Ubicazione",""),
            "IP": q.get("IP",""),
            "Generale": q.get("Interruttore generale (tipo/In)",""),
            "Diff": q.get("Differenziale generale (tipo/Idn, se presente)",""),
        })

    linee_list = []
    for _, r in linee_df_calc.iterrows():
        tipo = r.get("Tipo_cavo","")
        form = r.get("Formazione","")
        sez = r.get("Sezione_mm2","")
        cavo_str = f"{tipo} {form}x{sez} mm²" if tipo and form and sez else "XXXX (Inserire)"
        linee_list.append({
            "Linea": r.get("Circuito/Linea",""),
            "Uso": r.get("Destinazione/Utilizzo",""),
            "Posa": r.get("Posa",""),
            "L_m": r.get("Lunghezza_m",""),
            "Cavo": cavo_str,
            "Protezione": r.get("Protezione (MT/MTD)",""),
            "Diff": r.get("Differenziale (tipo/Idn)",""),
            "DV_perc": f"{r.get('ΔV_%','')}",
            "Esito": r.get("Esito",""),
        })

    payload = {
        "committente_nome": committente,
        "impianto_indirizzo": luogo,
        "data": data_doc.strftime("%d/%m/%Y"),
        "progettista_blocco": PROGETTISTA_BLOCCO,
        "premessa": premessa,
        "norme": norme,
        "dati_tecnici": dati_tecnici,
        "descrizione_impianto": descrizione_impianto,
        "confini": confini_txt,
        "quadri": quadri_list,
        "linee": linee_list,
        "sicurezza": sicurezza,
        "verifiche": verifiche,
        "manutenzione": manutenzione,
        "allegati": allegati,
        "firma": "Ing. Pasquale Senese",
        "oggetto_intervento": oggetto,
        "tipologia": tipologia,
        "sistema": sistema,
        "tensione": tensione,
        "potenza_disp": potenza_disp_kw,
        "n_doc": n_doc,
        "rev": revisione,
        "impresa": impresa,
    }

    pdf_bytes = genera_pdf_relazione_bytes(payload)
    st.success("PDF generato.")
    st.download_button(
        "Scarica PDF",
        data=pdf_bytes,
        file_name="Relazione_Tecnica_DiCo_Impianto_Elettrico.pdf",
        mime="application/pdf",
    )
