import streamlit as st
import pandas as pd
from datetime import date

from calcoli import corrente_da_potenza, caduta_tensione, verifica_tt_ra_idn, zs_massima_tn
from pdf_generator import genera_pdf_relazione_bytes

st.set_page_config(page_title="Relazione Tecnica DiCo – Impianti Elettrici", layout="wide")

st.title("Relazione Tecnica - Impianto Elettrico (Allegato alla DiCo)")
st.caption("Compilazione guidata (stile v7) + calcoli essenziali + generazione PDF.")

PROGETTISTA_BLOCCO = """Ing. Pasquale Senese
Via Francesco Soave 30 - 20135 Milano
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
    st.markdown("**Nota**: - ")

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
    cod_progetto = st.text_input("Cod. progetto", "XXXX (Inserire)")
    nome_progetto = st.text_input("Nome progetto", "XXXX (Inserire)")
    cover_style = st.selectbox("Stile cover", ["Engineering (title-block)", "A riquadri (legacy)"], index=0)

    n_doc = st.text_input("N. documento", "XXXX (Inserire)")
    revisione = st.text_input("Revisione", "00")
    data_doc = st.date_input("Data", value=date.today())

st.subheader("Revisioni documento")
rev_df = pd.DataFrame([
    {"Rev": str(revisione), "Data": data_doc.strftime('%d/%m/%Y'), "Descrizione": "Emissione documento"},
])
rev_df = st.data_editor(rev_df, num_rows="dynamic", use_container_width=True, key="revisioni")

# Carica opzionale immagine timbro/firma per la cover
timbro_file = st.file_uploader("Timbro/Firma (PNG) - opzionale", type=["png"], accept_multiple_files=False)
timbro_bytes = timbro_file.getvalue() if timbro_file else None

st.divider()

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
    st.markdown("**Progettista / Tecnico redattore**")
    progettista_blocco = st.text_area(
        "Dati progettista (blocco)",
        value=(
            "Ing. Pasquale Senese\n"
            "Via Francesco Soave 30 - 20135 Milano (MI) - Cell: 340 5731381\n"
            "Email: pasquale.senese@ingpec.eu  P.IVA: 14572980960"
        ),
        height=100,
    )

# Deriva il nominativo (prima riga) per cover/title-block
progettista_nome = (progettista_blocco.strip().splitlines()[0].strip() if progettista_blocco.strip() else "")

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

# Campi per eliminare XXXX in premessa/norme
st.subheader("Fonti dati e prescrizioni (per evitare 'XXXX' nel PDF)")
c1, c2 = st.columns(2)
with c1:
    fonte_dati = st.text_input("Fonte dati fornitura/condizioni (Committente/Impresa/Gestore)", "Committente")
with c2:
    prescrizioni_enti = st.text_input("Prescrizioni Enti/Autorità locali (se presenti)", "Nessuna / Non applicabile")


# =========================
# CRITERIO DI PROGETTO (ESTESO - da relazione tecnico-specialistica)
# =========================
st.subheader("Criterio di progetto degli impianti (testo esteso)")
st.caption(
    "Questo capitolo riprende la relazione tecnico-specialistica (con formule). "
    "Se lo disattivi, non verrà stampato nel PDF."
)
includi_criterio = st.checkbox("Includi capitolo 3 esteso (criterio di progetto)", value=True)
cosphi_ricarica = st.number_input(
    "Fattore di potenza (cosφ) per linee prese di ricarica (se presenti)",
    min_value=0.50, max_value=1.00, value=0.99, step=0.01
)
criterio_note = st.text_area("Note/integrazioni al capitolo 3 (opzionale)", "", height=90)


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
     "Differenziale (tipo/Idn)":"Tipo A 30mA", "Tipo_diff":"A", "Idn_mA":30,
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
        "Tipo_diff": st.column_config.SelectboxColumn("Tipo diff", options=["AC","A","F","B"], required=False),
        "Idn_mA": st.column_config.NumberColumn("Idn (mA)", min_value=0, max_value=3000, step=1),
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
        idn_a = float(row.get("Idn_mA") or 0.0) / 1000.0
        if ra > 0 and idn_a > 0:
            ok_tt = verifica_tt_ra_idn(ra, idn_a, ul=ul_tt)
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
# SICUREZZA / TERRA / SPD + CPI
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
    spd_quadro = st.text_input("Quadro di installazione SPD (se pertinente)", "")
    spd_caratt = st.text_input("Caratteristiche principali SPD (se pertinente)", "")

st.subheader("Prevenzione incendi / VV.F. (se pertinente)")
c1, c2, c3 = st.columns(3)
with c1:
    attivita_vvf = st.selectbox("Attività soggetta VV.F. (DPR 151/2011)", ["Non pertinente", "Sì", "No (da verificare)"], index=0)
with c2:
    cpi = st.selectbox("CPI / SCIA antincendio", ["Non pertinente", "Presente", "Non presente", "In corso"], index=0)
with c3:
    vvf_note = st.text_input("Note VV.F. (se pertinente)", "—")

st.divider()

# =========================
# VERIFICHE
# =========================
st.subheader("Verifiche, prove e collaudi (registro sintetico)")

ver_df = pd.DataFrame([
    {"Prova / Verifica":"Esame a vista", "Esito":"positivo", "Strumento":"—", "Note":"—"},
    {"Prova / Verifica":"Continuità PE ed equipotenziale", "Esito":"positivo", "Strumento":"—", "Note":"—"},
    {"Prova / Verifica":"Resistenza di isolamento", "Esito":"positivo", "Strumento":"—", "Note":"—"},
    {"Prova / Verifica":"Prova differenziali (Idn/tempo)", "Esito":"positivo", "Strumento":"—", "Note":"—"},
    {"Prova / Verifica":"Polarità / sequenza fasi (se pertinente)", "Esito":"non previsto", "Strumento":"—", "Note":"—"},
    {"Prova / Verifica":"TT: misura Ra e coordinamento con Idn (se TT)", "Esito":"positivo", "Strumento":"—", "Note":"—"},
    {"Prova / Verifica":"TN: misura Zs e verifica intervento (se TN)", "Esito":"non previsto", "Strumento":"—", "Note":"—"},
    {"Prova / Verifica":"Altre prove (SPD, emergenza, comandi, ecc.)", "Esito":"non previsto", "Strumento":"—", "Note":"—"},
])
ver_df = st.data_editor(ver_df, num_rows="dynamic", use_container_width=True, key="verifiche")

st.divider()

# =========================
# FIRMA
# =========================
st.subheader("Firma (stampa nel PDF)")
c1, c2, c3 = st.columns(3)
with c1:
    luogo_firma = st.text_input("Luogo firma", "XXXX (Inserire)")
with c2:
    data_firma = st.date_input("Data firma", value=data_doc)
with c3:
    firmatario = st.text_input("Firmatario", "Ing. Pasquale Senese")

st.divider()

# =========================
# GENERAZIONE PDF
# =========================
st.subheader("Genera PDF")

amb_txt = ", ".join([a for a in ambienti if a != "Altro"])
if "Altro" in ambienti:
    amb_txt += f", Altro: {amb_altro}"

premessa = f"""La presente Relazione Tecnico Specialistica è redatta nell’ambito dell’incarico conferito dalla Committenza "{committente}" e riguarda l’intervento "{oggetto}" presso "{luogo}".

FINALITÀ E PERIMETRO

Il documento ha lo scopo di:

• descrivere l’impianto e le opere eseguite/da eseguire, con indicazione dei confini dell’intervento;
• richiamare i riferimenti legislativi e normativi applicabili;
• esplicitare i criteri di progettazione e le verifiche di coordinamento essenziali (correnti, cadute di tensione, protezioni), in coerenza con la regola dell’arte.

Si precisa che la Stazione di ricarica non richiede la presenza di personale fisso e pertanto non va ad incidere in alcun modo sui parametri connessi con gli standard urbanistici e gli spazi ad essa relativa sono da ritenersi “senza permanenza di persone”.
In relazione a quanto previsto dalla vigente normativa in materia di portatori di handicap è opportuno precisare che i locali tecnici sono accessibili esclusivamente al personale specializzato per la manutenzione che non può essere svolta da persone con ridotte capacità motorie; in questo senso le prescrizioni di cui alla legge 09/01/1989 n° 13 e successive modificazioni, sono derogabili ai sensi dell’art. 7.4 del D.M. n° 236 del 14/06/1989.
Tutte le operazioni di accesso ai locali tecnici verranno effettuate in conformità al D.Lgs 81/2008 mediante utilizzo di idonea attrezzatura a norma.
Le infrastrutture di ricarica dei veicoli elettrici non rientrano fra le attività soggette ai controlli di prevenzione incendi ai sensi dell’Allegato I del D.P.R. n. 151 del 1 agosto 2011 “Regolamento recante semplificazione della disciplina dei procedimenti relativi alla prevenzione incendi, a norma dell’articolo 49 comma 4-quater, decreto-legge 31 maggio 2010, n. 78, convertito con modificazioni, dalla legge 30 luglio 2010, n. 122”, pertanto non sono soggette al benestare dei VV.F. e non necessitano di parere preventivo da parte del Comando Provinciale dei Vigili del Fuoco. 

Qualora l’installazione dell’infrastruttura di ricarica avvenga in una attività soggetta al controllo dei VV.F. si applicano le indicazioni riportate nelle linee guida contenute nella Circolare 2/2018 protocollo U.0015000 del 05/11/2018.

Saranno inoltre rispettate ed osservate le norme relative alla prevenzione infortuni ed alla sicurezza in cantiere.


VALENZA DOCUMENTALE

La presente Relazione costituisce documento tecnico di progetto e di supporto alla documentazione di conformità ai sensi del D.M. 37/2008; non sostituisce la Dichiarazione di Conformità (DiCo) né i relativi allegati obbligatori, che restano di competenza dell’Impresa installatrice.

RESPONSABILITÀ E DATI DI INGRESSO

Le informazioni relative alla fornitura elettrica (POD, potenza disponibile/contrattuale, caratteristiche del punto di consegna), destinazione d’uso e condizioni di esercizio sono state fornite da "{fonte_dati}" e/o rilevate in sito e/o confermate in data {data_doc.strftime('%d/%m/%Y')}. Eventuali porzioni preesistenti non oggetto di intervento e le interfacce con impianti/parti terze sono indicate nel paragrafo "Confini dell’intervento".

REQUISITI MATERIALI E CONSEGNA

Materiali e componenti devono essere conformi alle norme applicabili, provvisti di marcatura CE e, ove disponibile, marchio di conformità volontario (es. IMQ) o equivalente. Alla consegna l’impianto deve risultare conforme alla regola dell’arte e alle prescrizioni eventualmente impartite da Enti/Autorità competenti.
"""

norme = f"""Si riportano i principali riferimenti legislativi e normativi applicabili (elenco non esaustivo):

• D.M. 22/01/2008 n. 37.
• Legge 01/03/1968 n. 186.
• D.Lgs. 09/04/2008 n. 81 e s.m.i.
• D.P.R. 22/10/2001 n. 462 (ove applicabile).
• Norme CEI applicabili (in particolare CEI 64-8, CEI 64-14, CEI EN 61439, CEI EN 60529; e, se pertinenti, CEI 81-10, CEI 0-10, CEI 0-21/0-16).
• Regolamento Prodotti da Costruzione (UE) 305/2011 (CPR) e norme CEI-UNEL per i cavi (ove applicabile).

Eventuali ulteriori prescrizioni di Enti/Autorità locali: {prescrizioni_enti}.
"""

dati_tecnici = f"""Tipo sistema di distribuzione: {sistema}. Tensione nominale: {tensione}. Potenza disponibile/contrattuale: {potenza_disp_kw}.
POD: {pod} – contatore ubicato in: {contatore_ubi}.
Alimentazione: {alimentazione}. Potenza prevista/servita (stima): {potenza_prev_kw:.1f} kW (Ib indicativa ≈ {Ib:.1f} A a cosφ={cosphi:.2f}).
Ambientazioni particolari (se presenti): {amb_txt}.
"""

descrizione_impianto = f"""Il sito di intervento è ubicato in {luogo}. L’impianto è alimentato in bassa tensione dal punto di consegna del Distributore (POD: {pod}), tramite contatore/quadretto di misura ubicato in {contatore_ubi}.
Tipo sistema di distribuzione: {sistema}. Tensione nominale: {tensione}. Potenza disponibile/contrattuale: {potenza_disp_kw}.

La ripartizione e distribuzione interna avviene mediante linee in cavo conforme CEI/UNEL e componenti marcati CE (e, ove disponibile, IMQ o equivalente). Le condutture sono posate in tubazioni/canalizzazioni idonee e con protezione meccanica adeguata; i circuiti risultano identificati e separati per destinazione d’uso (illuminazione, prese, ausiliari, ecc.), privilegiando la manutenibilità.

Scopo dell’intervento (descrizione sintetica): {oggetto}

Le opere impiantistiche previste comprendono, in funzione dell’intervento, la realizzazione e/o modifica di linee di alimentazione dedicate, installazione di punti di utilizzo, posa di tubazioni/canalizzazioni, installazione o adeguamento di quadri elettrici (generale e/o di zona), apparecchi di protezione e comando, morsetterie e accessori, nonché collegamenti al sistema di protezione (PE) e ai collegamenti equipotenziali.

I conduttori sono identificati secondo codifica colori (PE giallo-verde, N blu, fasi marrone/nero/grigio) e marcatura/etichettatura dove previsto. I dispositivi di protezione sono coordinati con le linee e con il sistema di distribuzione (TT/TN) in modo coerente con le norme tecniche applicabili.
"""

confini_txt = f"""L’intervento comprende: {compresi}

Sono esclusi: {esclusi}

Integrazione con impianto esistente: {integrazione}. {("Descrizione e limiti: " + integrazione_note) if integrazione == "Sì" else ""}
"""

# Costruisci frase "Idn/tipo" a partire dai dati delle linee (se presenti)
diff_tipici = []
for _, r in linee_df_calc.iterrows():
    td = str(r.get("Tipo_diff") or "").strip()
    idn = int(r.get("Idn_mA") or 0)
    if td and idn:
        diff_tipici.append(f"Tipo {td} {idn} mA")
diff_frase = ", ".join(sorted(set(diff_tipici))) if diff_tipici else "N.D."

vvf_blocco = ""
if attivita_vvf != "Non pertinente" or cpi != "Non pertinente":
    vvf_blocco = f"Prevenzione incendi / VV.F.: attività soggetta: {attivita_vvf}; CPI/SCIA: {cpi}. Note: {vvf_note}."

sicurezza = f"""La protezione contro i contatti diretti è assicurata tramite isolamento delle parti attive, involucri/barriere con grado di protezione adeguato e corretta posa delle condutture.

La protezione contro i contatti indiretti è assicurata mediante interruzione automatica dell’alimentazione, in accordo con CEI 64-8, tramite dispositivi differenziali e/o magnetotermici coordinati con l’impianto di terra (nei sistemi TT) o con il conduttore di protezione (nei sistemi TN).

Protezione differenziale adottata (sintesi): {diff_frase}.

Configurazione impianto di terra: {terra_cfg}. Dispersore: {dispersore}. Collegamenti equipotenziali principali: {equipot}.

Protezione contro le sovratensioni (SPD) – esito: {spd_esito}. {("Tipologia: " + ", ".join(spd_tipo) + " – ") if spd_tipo else ""}quadro: {spd_quadro}. Caratteristiche: {spd_caratt}.

Caduta di tensione: verificata entro il limite adottato in progetto: {dv_lim:.1f}%.
{vvf_blocco}
"""

verifiche = """Ad ultimazione dei lavori, l’impianto è sottoposto alle verifiche previste dalla CEI 64-8 (Parte 6) e dalla CEI 64-14, con esecuzione e registrazione delle prove strumentali pertinenti al sistema di distribuzione (TT/TN) e alla tipologia di impianto. In particolare:\n\n"""
for _, r in ver_df.iterrows():
    verifiche += f"• {r.get('Prova / Verifica','')}: {r.get('Esito','')} – Strumento: {r.get('Strumento','')} – Note: {r.get('Note','')}\n"

manutenzione = """Le attività di esercizio e manutenzione devono essere svolte da personale qualificato e autorizzato, in sicurezza e nel rispetto delle istruzioni dei costruttori e delle norme tecniche applicabili (es. CEI 0-10 / CEI 11-27, ove pertinenti).

PIANO DI MANUTENZIONE (minimo consigliato)

• Quadri elettrici: ispezione visiva, pulizia, verifica serraggi morsetti, integrità targhe/etichette e dispositivi di protezione;
• Dispositivi differenziali: prova periodica con tasto "T" e verifiche strumentali (Idn/tempo) secondo periodicità e criticità del sito;
• Conduttori e condutture: verifica integrità isolamento, fissaggi, protezioni meccaniche e segregazioni;
• Collegamenti equipotenziali e PE: controllo continuità e integrità;
• Comandi/emergenze (se presenti): prova funzionale e ripristino, verifica segnalazioni e cartellonistica;
• Apparecchiature specifiche (es. wallbox/utenze dedicate): ispezione cavi e connettori, prova funzionale e aggiornamenti firmware se previsti dal costruttore.

È raccomandata la tenuta di un registro manutenzione con data, attività eseguite, esito e nominativo dell’operatore."""

allegati = """Completano la presente relazione e/o la DiCo i seguenti allegati. 

- Schema unifilare / multifilare dei quadri interessati: Obbligatorio.
- Elenco linee/circuiti con cavo e protezione (tabella circuiti): Obbligatorio.
- Verbali e report delle misure e prove strumentali
- Schede tecniche principali componenti (quadri, interruttori, SPD, ecc.): Se disponibile.
- Dichiarazioni/Marcature CE (ed eventuale IMQ) dei materiali: Se disponibile.
- Report fotografico essenziale (quadri, targhette, collegamenti di terra, punti significativi): Consigliato.
"""

# =========================
# ALLEGATI FOTOGRAFICI (4 FOTO IN UN'UNICA PAGINA)
# =========================
st.subheader("Allegati fotografici (opzionali)")
st.caption("Carica fino a 4 foto: verranno inserite nel PDF in un'unica pagina (griglia 2×2).")

fc1, fc2 = st.columns(2)
with fc1:
    foto1_file = st.file_uploader(
        "Foto 1 – Posizione Pulsante Antincendio (se presente)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=False,
        key="foto1",
    )
    if foto1_file:
        st.image(foto1_file, caption="Foto 1", use_container_width=True)

    foto3_file = st.file_uploader(
        "Foto 3 – Percorso realizzato",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=False,
        key="foto3",
    )
    if foto3_file:
        st.image(foto3_file, caption="Foto 3", use_container_width=True)

with fc2:
    foto2_file = st.file_uploader(
        "Foto 2 – Quadro realizzato",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=False,
        key="foto2",
    )
    if foto2_file:
        st.image(foto2_file, caption="Foto 2", use_container_width=True)

    foto4_file = st.file_uploader(
        "Foto 4 – Apparecchiatura di Ricarica (se installata)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=False,
        key="foto4",
    )
    if foto4_file:
        st.image(foto4_file, caption="Foto 4", use_container_width=True)

if st.button("Genera PDF"):
    foto1_bytes = foto1_file.getvalue() if foto1_file else None
    foto2_bytes = foto2_file.getvalue() if foto2_file else None
    foto3_bytes = foto3_file.getvalue() if foto3_file else None
    foto4_bytes = foto4_file.getvalue() if foto4_file else None

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
        cavo_str = f"{tipo} {form}x{sez} mm²" if tipo and form and sez else ""
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


    # === CAPITOLO 3 - CRITERIO DI PROGETTO (ESTESO) ===
    criterio_testo = ""
    if includi_criterio:
        # Tipi cavo usati nelle linee (se presenti)
        try:
            tipi_cavo_usati = ", ".join(sorted(set([str(x) for x in linee_df_calc.get("Tipo_cavo", []).dropna().tolist()])))
        except Exception:
            tipi_cavo_usati = "N.D."

        criterio_testo = (
f"""Tutti i materiali e le apparecchiature utilizzati devono essere di alta qualità, prodotti da aziende affidabili, ben lavorati e adatti all'uso previsto, resistendo a sollecitazioni meccaniche, corrosione, calore, umidità e acque meteoriche (per installazione all’esterno). Devono garantire lunga durata, facilità di ispezione e manutenzione.
È obbligatorio l'uso di componenti con marcatura CE e, se disponibile, marchio IMQ o equivalente europeo. I componenti senza marcatura CE devono avere una dichiarazione di conformità del costruttore ai requisiti di sicurezza delle normative CEI, UNI o IEC.

3.1 Dimensionamento delle linee

Le linee elettriche sono calcolate mediante l’utilizzo dei seguenti criteri progettuali:
• La corrente di impiego (Ib) è calcolata considerando la potenza nominale delle apparecchiature elettriche. La tensione di alimentazione è pari a 230 V per le utenze monofase, 400 V per le utenze trifase. Fattore di potenza pari a {cosphi_ricarica:.2f} per le linee di alimentazione delle prese di ricarica (se presenti).
• La corrente nominale della protezione (In), definita dal costruttore, è considerata come la corrente che l’interruttore può sopportare per un tempo indefinito senza che quest’ultimo subisca alcun danno.
• La portata del cavo (Iz) è calcolata utilizzando le tabelle CEI UNEL 35024 e 35026, tenendo conto delle condizioni di posa, del tipo di isolante del cavo e della temperatura ambiente.
I cavi di alimentazione sono dimensionati in modo da non subire danneggiamento causato da sovraccarichi e cortocircuiti mediante il coordinamento con la corrente nominale (In) del dispositivo di protezione a monte (vedi paragrafi 3.5.1 e 3.5.2).

3.2 Calcolo della sezione del cavo in funzione della corrente di impiego (Ib)

Nota la potenza assorbita dall’utenza, la corrente d’impiego (Ib) può essere calcolata come:
Ib = (Ku · P) / (k · Vn · cosφ)

dove:
• k = 1 per i circuiti monofase; k = √3 per i circuiti trifase;
• Ku è il coefficiente di utilizzazione della potenza nominale del carico;
• P è la potenza totale dell’utenza [W];
• Vn è la tensione nominale del sistema [V].
Determinata la corrente di impiego per ogni utenza, è possibile dimensionare il cavo con portata Iz > Ib.

3.3 Caduta di tensione

Dopo aver determinato la sezione del cavo in funzione della corrente d’impiego, si verifica la caduta di tensione con la formula:
ΔV = K · (R·cosφ + X·sinφ) · L · I

dove:
• K = 2 per le linee monofase (230 V); K = √3 per le linee trifase (400 V);
• R e X sono resistenza e reattanza per unità di lunghezza [Ω/km];
• I è la corrente di impiego;
• L è la lunghezza della linea [m].
La caduta di tensione percentuale è:
ΔV% = (ΔV / Vn) · 100
La caduta di tensione percentuale complessiva non deve superare {dv_lim:.1f}% (rif. CEI 64-8 art. 525).

3.4 Sezione e tipologia dei cavi utilizzati

I cavi utilizzati sono conformi al Regolamento UE 305/2011 (CPR), all’unificazione UNEL e alle norme costruttive CEI.
Per il dimensionamento dei conduttori di neutro e del conduttore di protezione (PE) si fa riferimento alla CEI 64-8/5 par. 543.1.2 tabella 54F:
• per sezione fase Sf ≤ 16 mm²: SPE = Sf
• per 16 < Sf ≤ 35 mm²: SPE = 16 mm²
• per Sf > 35 mm²: SPE = Sf/2
Qualora il PE non faccia parte della conduttura di alimentazione (CEI 64-8/5 par. 543.1.3), valgono i criteri sopra con minimi: 2,5 mm² Cu (con protezione meccanica) o 4 mm² Cu (senza protezione meccanica).

3.4.1 Tipologia dei cavi

Quanto segue è valido solo nel caso in cui il POD, e quindi il contatore fiscale ed il DG non sia all’interno di EVC, come nel caso oggetto della presente relazione tecnica.
Tutti i cavi impiegati nella realizzazione degli impianti descritti nel presente progetto risponderanno all'unificazione UNEL ed alle Norme costruttive stabilite dal Comitato Elettrotecnico Italiano.
In particolare, tutti i cavi citati nel presente documento si intendono del tipo non propagante l’incendio ed a bassissima emissione di fumi e gas tossici, secondo le norme vigenti quali CEI 20-22 III, CEI 20-35, CEI 20-37 e CEI 20-38 o s.m.i.
Per la distribuzione dell’energia dovranno essere utilizzati i cavi unipolari isolati in XLPE, qualità R2 non propaganti l’incendio, con corde flessibili in rame, rispondenti alla norma CEI 20-22, ed avranno una tensione di isolamento minimo, superiore di un gradino alla tensione di impiego (Uo/U = 0,6/1 kV).
I cavi saranno contrassegnati in modo da individuare prontamente il servizio a cui appartengono; il transito di cavi attraverso la struttura di canali portacavi, cassette di derivazione etc., sarà effettuato con l'ausilio di pressacavi del tipo con bullone a stringere.
I conduttori previsti saranno dimensionati secondo i dati della tabella CEI-UNEL 35024/1 e 35024/2 tenendo conto di una temperatura iniziale di 30°C, di una temperatura massima di esercizio e di una temperatura massima di corto circuito adeguati al tipo dell'isolante (CEI 64-8 tabella 52 D); per la posa interrata si farà riferimento alle tabelle CEI-UNEL 35026.
Nel caso siano posati nella stessa conduttura conduttori di sistemi a tensione diversa (cavi per energia, impianto rivelazione incendio, impianti trasmissione dati, ecc.), tutti i conduttori dovranno essere isolati per la tensione più elevata (CEI 64-8 art. 521.6).

I cavi impiegati nel progetto (in funzione delle tratte e delle modalità di posa) appartengono alle tipologie selezionate nei circuiti: {tipi_cavo_usati}.
Esempi (se pertinenti):
• FG16(M)16 / FG16(O)M16 (o similari) per dorsali/esterni Uo/U 0,6/1 kV (HEPR G16 + guaina R16) – CEI UNEL 35318/35322.
• FS17 450/750 V per cablaggi interni quadro e PE (unipolare senza guaina, PVC S17, CPR).

In accordo con la Tabella 52A della Norma CEI 64-8, si potranno utilizzare, ad esempio, i seguenti tipi di cavo:
•	posa all’interno e all’esterno non interrata: H07V-K, FS17, FG17 – 450/750 V;
•	posa all’interno e all’esterno anche interrata: FG16OR16-0,6/1 kV, FG16R16-0,6/1 kV, N1VV-K.
Per gli ambienti trattati nella Sezione 751 della Norma CEI 64-8: 
•	FM9 450/750 V
•	FG10(O)M1-0,6/1 kV.
Più in dettaglio, nel presente progetto si potranno usare i seguenti:
•	cavi unipolari o multipolari isolati in gomma CPR ad alto modulo, non propaganti l’incendio e a bassissima emissione di gas alogenidrici, con conduttori in rame ricotto, tensione nominale di 0,6/1 kV sigla di riferimento FG16(O)M16;
•	cavi unipolari o multipolari isolati in gomma CPR ad alto modulo, non propaganti l’incendio con conduttori in rame ricotto, tensione nominale di 0,6/1 kV sigla di riferimento FG16R16;
•	cavi multipolari isolati in PVC, qualità TI2, non propaganti l’incendio con conduttori in rame ricotto, tensione nominale 450/750 V sigla di riferimento FG17;
•	cavi unipolari isolati in PVC, qualità R2 non propaganti l’incendio, con corde flessibili in rame, per tensioni nominali 450/750 V sigla di riferimento FS17.


3.4.2 Posa dei cavi

Le tipologie di posa sono indicate nella tabella circuiti (campo “Posa”) e possono comprendere: tubazioni incassate/esterne, canalizzazioni, passerelle, tubazioni interrate, ecc. Gli attraversamenti di pareti/solai saranno ripristinati, ove necessario, con sigillature idonee a mantenere la compartimentazione. 
Nei punti in cui le condutture e/o le tubazioni impiantistiche attraversano elementi di separazione resistenti al fuoco (pareti e solai di compartimentazione), dovrà essere garantito il mantenimento della prestazione di compartimentazione prevista dal progetto antincendio. In conformità ai principi del Codice di Prevenzione Incendi (D.M. 03/08/2015 e s.m.i.) e alle norme di prova e classificazione della resistenza al fuoco dei sistemi di attraversamento, tutti i fori e i passaggi dovranno essere ripristinati mediante sistemi di sigillatura certificati (firestop) con classificazione almeno pari a quella dell’elemento attraversato (es. EI/REI richiesto), installati secondo le istruzioni del produttore.
A titolo esemplificativo, per tubazioni combustibili (PVC, PE, PP – tipicamente scarichi e pluviali) si impiegheranno collari tagliafuoco/REI con materiale termoespandente (intumescente) che, in caso d’incendio, occlude il foro sigillando il passaggio; per cavidotti/cavi e canalizzazioni si utilizzeranno idonei sistemi (malte o sigillanti intumescenti, schiume certificate, bende/manicotti, pannelli o cuscini) compatibili con il tipo di impianto e con le condizioni di posa.
I prodotti utilizzati dovranno essere marcati CE ove applicabile ai sensi del Regolamento (UE) 305/2011 (CPR) oppure corredati da Valutazione Tecnica Europea (ETA) e Dichiarazione di Prestazione (DoP), con rapporti di prova secondo UNI EN 1366-3 (sigillature di attraversamenti) e classificazione secondo UNI EN 13501-2. L’impresa incaricata dell’esecuzione degli attraversamenti e del ripristino dovrà impiegare materiali idonei e certificati, assicurare la continuità della tenuta ai fumi e ai gas caldi e rilasciare idonea documentazione di posa (schede prodotto, istruzioni e, ove richiesto, dichiarazione di corretta installazione) a garanzia del mantenimento della compartimentazione di progetto

Le condutture elettriche dovranno essere opportunamente distanziate da tubazioni che producano calore, fumi o vapori. Se ciò non fosse possibile si dovranno utilizzare opportuni accorgimenti onde evitare eventuali effetti dannosi.
I tubi protettivi installati sotto traccia dovranno avere un percorso orizzontale, verticale o parallelo allo spigolo della parete, ad esclusione dei percorsi nei soffitti e nei pavimenti ove il percorso potrà essere omnidirezionale.
I tipi di posa delle condutture in funzione dei tipi di cavi utilizzati dovranno essere in accordo con la Tabella 52A della norma CEI 64-8. Per i conduttori nudi la posa è permessa su isolatori, mentre non è prevista nei tubi protettivi di forma non circolare, su passerelle e mensole e con filo o corda di supporto. Per i cavi senza guaina la posa è permessa nei tubi protettivi di forma non circolare e su passerelle e mensole, mentre non è permessa con filo o corda di supporto. Per i cavi con guaina multipolari la posa è permessa nei tubi protettivi di forma non circolare, su passerelle e mensole e con filo o corda di supporto, mentre su isolatori non è applicabile o non è usata in generale nella pratica. Per i cavi con guaina unipolari la posa è permessa nei tubi protettivi di forma non circolare, su passerelle e mensole e con filo o corda di supporto, mentre su isolatori non è applicabile o non è usata in generale nella pratica.
Sempre in accordo con la Tabella 52A, per i cavi con guaina multipolari la posa senza fissaggi, con fissaggio diretto su parete, entro tubi protettivi di forma circolare e in canali (compresi i canali incassati nel pavimento) è indicata nella tabella con combinazioni di casi non permessi e casi permessi; in particolare, per i cavi con guaina multipolari sono riportati casi non permessi e casi permessi per la posa in tubi protettivi di forma circolare e in canali, oltre a casi permessi per la posa con fissaggio diretto su parete. Per i cavi con guaina unipolari la posa senza fissaggi è non applicabile o non usata in generale nella pratica, mentre è permessa con fissaggio diretto su parete, entro tubi protettivi di forma circolare e in canali (compresi i canali incassati nel pavimento).
Nei cavi con guaina sono compresi i cavi provvisti di armatura e quelli con isolamento minerale. La legenda della tabella è la seguente: il simbolo “+” indica “permesso”, il simbolo “-” indica “non permesso” e il simbolo “0” indica “non applicabile o non usato in generale nella pratica”.
I tipi di posa delle condutture in funzione delle varie condizioni di utilizzo dovranno essere in accordo con la Tabella 52B della norma CEI 64-8. Entro cavità di strutture accessibili sono previsti i codici 25 per la posa senza fissaggi, 21-25 per la posa con fissaggio diretto su parete, 22 per la posa entro tubi protettivi di forma circolare e 31-32-75 per la posa in canali, compresi i canali incassati nel pavimento. Entro cavità di strutture non accessibili sono previsti i codici 21-25-73-74 per la posa senza fissaggi, 0 per la posa con fissaggio diretto su parete, 2-73-74 per la posa entro tubi protettivi di forma circolare e 0 per la posa in canali. Entro cunicoli sono previsti i codici 43 per la posa senza fissaggi, 43 per la posa con fissaggio diretto su parete, 41-42 per la posa entro tubi protettivi di forma circolare e 0 per la posa in canali. Per la posa interrata sono previsti i codici 62-63 per la posa senza fissaggi, 0 per la posa con fissaggio diretto su parete, 61 per la posa entro tubi protettivi di forma circolare e “-” per la posa in canali. Per la posa incassata nella struttura sono previsti i codici 52-53 per la posa senza fissaggi, 51 per la posa con fissaggio diretto su parete, 1-2-5 per la posa entro tubi protettivi di forma circolare e 33-75 per la posa in canali. Per il montaggio sporgente sono previsti “-” per la posa senza fissaggi, 11 per la posa con fissaggio diretto su parete, 3 per la posa entro tubi protettivi di forma circolare e 31-32-71-72 per la posa in canali. Per la posa aerea sono previsti “-” per la posa senza fissaggi, “-” per la posa con fissaggio diretto su parete, 0 per la posa entro tubi protettivi di forma circolare e 34 per la posa in canali. Per la posa immersa sono previsti i codici 81 per la posa senza fissaggi, 81 per la posa con fissaggio diretto su parete, 0 per la posa entro tubi protettivi di forma circolare e “-” per la posa in canali.

Le condutture elettriche dovranno essere opportunamente distanziate da tubazioni che producano calore, fumi o vapori. Se ciò non fosse possibile si dovranno utilizzare opportuni accorgimenti onde evitare eventuali effetti dannosi.
I tubi protettivi installati sotto traccia dovranno avere un percorso orizzontale, verticale o parallelo allo spigolo della parete, ad esclusione dei percorsi nei soffitti e nei pavimenti ove il percorso potrà essere omnidirezionale.


Nei punti in cui le condutture e/o le tubazioni impiantistiche attraversano elementi di separazione resistenti al fuoco (pareti e solai di compartimentazione), dovrà essere garantito il mantenimento della prestazione di compartimentazione prevista dal progetto antincendio. In conformità ai principi del Codice di Prevenzione Incendi (D.M. 03/08/2015 e s.m.i.) e alle norme di prova e classificazione della resistenza al fuoco dei sistemi di attraversamento, tutti i fori e i passaggi dovranno essere ripristinati mediante sistemi di sigillatura certificati (firestop) con classificazione almeno pari a quella dell’elemento attraversato (es. EI/REI richiesto), installati secondo le istruzioni del produttore.
I prodotti utilizzati dovranno essere marcati CE ove applicabile ai sensi del Regolamento (UE) 305/2011 (CPR) oppure corredati da Valutazione Tecnica Europea (ETA) e Dichiarazione di Prestazione (DoP), con rapporti di prova secondo UNI EN 1366-3 e classificazione secondo UNI EN 13501-2.
Documentazione minima obbligatoria (a tutela della compartimentazione): l’Impresa incaricata dell’esecuzione degli attraversamenti e del ripristino dovrà consegnare, per ciascun attraversamento, registro attraversamenti (identificativo, ubicazione, elemento attraversato, prestazione richiesta, sistema adottato), schede prodotto/ETA/DoP, istruzioni di posa, e dichiarazione di corretta installazione, corredando il tutto con documentazione fotografica prima/dopo.
In mancanza della suddetta documentazione, la verifica della conformità delle sigillature firestop non si intende effettuata dal Progettista/Tecnico redattore e resta in capo all’Impresa e alla Direzione Lavori/Committente secondo le rispettive competenze.
Nota: Il presente documento non costituisce progetto antincendio né asseverazione ai fini della prevenzione incendi; i requisiti EI/REI e le soluzioni di compartimentazione sono quelli definiti dal progetto antincendio e dalle relative certificazioni.

3.4.3 Colorazione dei conduttori

I conduttori impiegati nell’esecuzione degli impianti dovranno essere contraddistinti dalla colorazione prevista dalle vigenti tabelle di unificazione CEI - UNEL 00722 e 00712.
Per quanto riguarda i conduttori di fase dovranno essere contraddistinti in modo univoco per tutto l'impianto dai colori: nero, grigio e marrone. Nella scelta del colore dei conduttori, il bicolore giallo-verde sarà tassativamente riservato ai conduttori di protezione ed equipotenziali ed il colore blu chiaro sarà destinato esclusivamente al conduttore di neutro (CEI 64-8 art. 514.3.1).

I conduttori sono identificati secondo CEI-UNEL 00722 e 00712:
• PE: giallo/verde; • Neutro: blu; • Fasi: marrone/grigio/nero.

3.4.4 Condutture

Quanto segue è valido solo nel caso in cui il POD, e quindi il contatore fiscale ed il DG non sia all’interno di EVC, come nel caso oggetto della presente relazione tecnica.
Le condutture dovranno essere realizzate in modo da ridurre al minimo la probabilità di innesco e propagazione dell’incendio nelle condizioni di posa. Per soddisfare questi requisiti le condutture dovranno rispondere alle prescrizioni della Sezione 751 della Norma CEI 64-8/7.
Per conduttura si dovrà intendere l'insieme costituito da uno o più conduttori elettrici e dagli elementi che assicurano il loro isolamento, il loro supporto, il loro fissaggio e la loro eventuale protezione meccanica (CEI 64-8/2 art. 26.1).
I conduttori dovranno essere sempre protetti meccanicamente. Dette protezioni saranno realizzate mediante tubazioni anche interrate, canalette portacavi, passerelle, condotti o cunicoli, eventualmente ricavati nella struttura edile ecc.
I tubi protettivi, le cassette e le scatole per l'impianto di energia, per trasmissione dati, di allarme, di controllo e di segnalazione, dovranno essere dedicate e distinte fra loro (CEI 64-8/5 art. 528.1.1).
Le condutture elettriche dovranno essere opportunamente distanziate da tubazioni che producano calore, fumi o vapori. Se ciò non fosse possibile si dovranno utilizzare opportuni accorgimenti onde evitare eventuali effetti dannosi.

3.4.5 Tubi e Guaine

In considerazione delle diverse tipologie impiantistiche si potranno utilizzare, oltre a quelli già esistenti, i tubi e le guaine di seguito descritte:
•	tubo rigido autoestinguente in PVC serie pesante conforme alla Norma CEI 23-8 e varianti ed alle relative tabelle UNEL 37118-37119-37120 e s.m.i.;
•	tubo flessibile autoestinguente in PVC serie pesante conforme alla Norma CEI 23-14 e varianti;
•	guaine in PVC flessibile autoestinguente, serie pesante, complete di accessori di giunzione e derivazione, conformi alle relative tabelle UNEL 37118-37119-37120 e s.m.i.
Il diametro dei tubi non dovrà essere inferiore a 16 mm. Tutte le curve eseguite senza l’impiego di pezzi speciali dovranno essere di raggio proporzionato al diametro del tubo e tale da non diminuirne in corrispondenza delle stesse la sezione libera di passaggio.
I tubi di nuova installazione dovranno essere dimensionati in modo che il loro diametro sia pari ad almeno 1,3 volte il diametro del cerchio circoscritto al fascio dei conduttori in essi contenuti.
Tale accorgimento renderà possibile un'eventuale aggiunta di conduttori senza arrecare deterioramento all'isolamento degli esistenti e permetterà di non apportare pregiudizio alla sfilabilità dei cavi.
Tutte le tubazioni, qualunque sia il tipo di posa, dovranno avere andamento prevalentemente rettilineo, si potranno seguire percorsi non rigorosamente rettilinei solamente in corrispondenza di eventuali ostacoli (canali, tubazioni di altri impianti).

3.4.6 Colonnina 

La stazione di ricarica è classificabile come “Ambienti ed applicazioni particolari, Alimentazione di veicoli elettrici” della Norma CEI 64-8:2012:06, variante V1:2013:07, Sezione 722.
Il modo di ricarica sarà tipo 3 e 4, mentre il modo di connessione sarà di tipo B e C (Norma CEI EN 61851-1:2012-05).
Attualmente la norma, che riporta le prescrizioni necessarie per la ricarica dei veicoli elettrici, è la Norma CEI EN 61851-1:2012-05 “Sistema di ricarica conduttiva dei veicoli elettrici – Parte 1: Prescrizioni generali
“…con riferimento ai modi di carica in corrente alternata adottati in Italia, al fine di garantire la necessaria sicurezza durante la carica conduttiva dei veicoli elettrici, quando questa viene eseguita in ambienti aperti a terzi deve essere adottato il Modo di carica 3”.

Sulla base delle classificazioni realizzate da Cives ed Eurelectric, il Piano Nazionale individua le seguenti classi di infrastrutture di ricarica sulla base della capacità di erogazione dell’energia:
•	Normal power (Slow charging) - fino a 3,7 kW
•	Medium power (Quick charging) - da 3,7 fino a 22 kW
•	High power (Fast charging) - superiore a 22 kW


3.5 Protezioni dalle sovracorrenti

La protezione dalle sovracorrenti è assicurata da interruttori automatici magnetotermici dimensionati affinché le curve I–t si mantengano al di sotto delle curve dei cavi protetti. Gli interruttori devono:
• interrompere sovraccarichi e cortocircuiti prima di danni all’isolamento;
• essere installati all’origine di ogni circuito/derivazione con portate differenti;
• avere PdI/Icu > Icc presunta nel punto di installazione.

3.5.1 Sovraccarichi (CEI 64-8 art. 433.2)
Tutte le condutture saranno protette dai sovraccarichi, con la sola esclusione dei circuiti la cui interruzione potrebbe dar luogo a pericolo per le persone. Le protezioni dai sovraccarichi saranno realizzate con interruttori automatici, rispondenti alle norme CEI 17-5 e CEI 23-3.
Ib ≤ In ≤ Iz
If ≤ 1,45 · Iz
dove:

•	In è la corrente nominale dell’interruttore o la sua taratura termica;
•	If è la corrente convenzionale di funzionamento dell’interruttore;
•	Ib è la corrente d’impiego;
•	Iz è la portata della linea.
Per quanto riguarda il soddisfacimento della seconda condizione, si terrà presente che:

•	gli interruttori per uso domestico o similare (norma CEI 23-3 e 23-18) hanno una corrente di funzionamento If = 1,45 x In;
•	gli interruttori conformi alla norma CEI 17-5 hanno una corrente di funzionamento If = 1,35 x In, per correnti nominali fino a 63 A e If = 1,25 x In, per valori della corrente nominale superiori a 63 A.
•	Quando la protezione dalle sovracorrenti sarà effettuata con fusibili si terranno presenti le relazioni in normativa applicabile.
3.5.2 Cortocircuiti (CEI 64-8 art. 434.3)
Per la protezione da corto circuito (CEI 64-8 art. 434.3), affinché la temperatura dei conduttori non superi il valore massimo ammissibile, si dovrà tener conto della relazione seguente:
I² · t ≤ K² · S²
dove:
•	I = corrente di corto circuito in Ampere;
•	t = durata del corto circuito in secondi;
•	K = fattore relativo alla natura dell'isolante
-	115 per cavo in rame con guaina esterna in PVC;
-	135 per cavi in rame isolati con gomma ordinaria o gomma butilica;
-	143 per cavi in rame isolati con gomma etilenpropilenica e propilene reticolato.
•	S = sezione del conduttore in mm.

3.6 Protezione dai contatti indiretti

La protezione contro i contatti indiretti è realizzata mediante interruzione automatica dell’alimentazione (TT/TN) e/o componenti a doppio isolamento.
Per la protezione dai contatti indiretti dovrà essere garantito il coordinamento dell’impianto di terra con i dispositivi di protezione (CEI 64-8/4 art. 413.1.4.2) in modo da assicurare l'interruzione automatica dell'alimentazione nei tempi richiesti.
Il coordinamento sarà soddisfatto dalla relazione:

Ra * Ia < 50

dove:
•	Ra = somma della resistenza del dispersore e dei conduttori di protezione
•	Ia = corrente che provoca il funzionamento automatico del dispositivo di protezione (Idn se il dispositivo è differenziale).
Nel caso di dispositivo con caratteristica di funzionamento a tempo inverso (interruttore magnetotermico) si dovrà garantire che tra una parte attiva e una massa (o un conduttore di protezione) non possa permanere una tensione di contatto superiore a 50 V (in corrente alternata) per un tempo superiore a 5s (CEI 64-8 art. 413.1.4.2).
Nell'utilizzo di dispositivi differenziali, che dovranno rispettare le prescrizioni della Norma CEI 23-18, l'intervento dovrà essere istantaneo.
Se si usano dispositivi differenziali di tipo selettivo (S) o ritardati, posti in serie a dispositivi differenziali di tipo generale, il tempo di intervento non dovrà essere superiore a 1s.


3.6.1 Sistema TT (CEI 64-8 art. 413.1.4.2)

Idn ≤ UL / Rt
con UL = {ul_tt:.0f} V (ambiente ordinario) e Rt resistenza complessiva terra+conduttori di protezione.

3.7 Protezione dai contatti diretti (CEI 64-8 art. 412)
Le parti attive risultano ricoperte tramite un isolamento che può essere rimosso solo mediante distruzione, Art. 412.1 CEI 64-8/4.
Ne consegue che i componenti in tensione e le parti attive dovranno essere segregati, mediante posa entro involucri o dietro barriere, in modo da assicurare un grado di protezione IPXXB (CEI 64-8 art. 412.2.1).
Per le superfici superiori orizzontali degli involucri e delle barriere a portata di mano si dovrà garantire un grado di protezione IPXXD (CEI 64-8 art. 412.2.2).
Nei luoghi soggetti a normativa specifica o con ambienti ed applicazioni particolari, il grado di protezione dovrà essere adeguato ai singoli casi, considerati in dettaglio nei capitoli specifici.
Le barriere e/o gli involucri di protezione dovranno essere fissati saldamente in modo da garantire stabilità e durata nel tempo e dovranno poter essere rimossi esclusivamente:
•	mediante l'uso di chiave o attrezzo;
•	se l'alimentazione, dopo l'interruzione a seguito della rimozione degli involucri di protezione, sia ripristinabile solo con la richiusura degli stessi;
•	se esiste una barriera intermedia, con grado di protezione minimo IPXXB, rimovibile solo con l'uso di chiave od attrezzo.
Sono possibili altri sistemi di protezione dai contatti diretti (ostacoli, distanziamento ecc.) che dovranno in ogni modo essere analizzati ed applicati solo in casi particolari e specifici (CEI 64-8 art. 412.2.4).

Isolamento delle parti attive e/o involucri/barriere (minimo IPXXB; superfici orizzontali a portata di mano: IPXXD). Vernici/smalti da soli non sono idonei.

3.8 Potere di interruzione delle apparecchiature

Icc-max < PdI (Icu) del dispositivo di protezione (rif. CEI EN 60947-2).

3.9 Quadri elettrici
Quadri conformi a CEI EN 61439-1/2 (e/o CEI 23-51 per domestici/similari). Cablaggio interno con conduttori idonei (es. FS17 CPR) e dimensionato per corrente nominale e cortocircuito nel punto di installazione.
Tutti i quadri e i dispositivi di protezione scelti (fusibili, interruttori) saranno di primaria marca: e dovranno essere del tipo in materiale termoplastico salvo diversa indicazione.
I quadri saranno completi di telai e pannellature idonee per il montaggio di apparecchi modulari e scatolati, e dovranno essere corredati di appositi cartellini fissati in modo imperdibile che indicheranno chiaramente le funzioni svolte dalle varie apparecchiature installate.
Sui quadri troveranno posto le protezioni magnetotermiche differenziali necessarie per attuare la protezione, il sezionamento e la suddivisione dei circuiti previsti con riferimento alla vigente normativa e in considerazione delle esigenze di sicurezza, continuità del servizio e praticità di manutenzione.
L'ingresso delle condutture (cavi provenienti dal contatore) sarà realizzato nella parte inferiore dello stesso; analogamente l’uscita delle condutture (cavi verso le utenze) sarà realizzata nella parte inferiore dello stesso.
Si raccomanda, nell’ingresso delle condutture al quadro, il mantenimento del grado di protezione iniziale dello stesso, mediante l’utilizzo di appositi pressa-cavi o guarnizioni. Le dimensioni dei quadri e le caratteristiche tecniche delle apparecchiature in essi installate sono specificate negli schemi elettrici allegati.
3.10 Distribuzione Generale 
Quanto segue è valido solo nel caso in cui il POD, e quindi il contatore fiscale ed il DG.
Gli ingressi delle condutture dall’Ente Distributore e delle condutture del fornitore dei servizi sono convogliati alla conchiglia dei quadri a mezzo di tubazione interrata all’interno della proprietà.
Le condutture della distribuzione saranno unicamente cavi unipolari tipo FG16R16/FG16M16, che saranno usati per sezioni di conduttore superiore a 25 mm2, mentre per sezioni inferiore sarà ammesso uso di cavi multipolare del tipo FG16(O)R16 o FRG17, che saranno utilizzati anche per le alimentazioni di utenze o i collegamenti di segnale nei locali tecnologici.

{("Integrazioni: " + criterio_note) if criterio_note.strip() else ""}"""
        )

    payload = {
        "committente_nome": committente,
        "impianto_indirizzo": luogo,
        "data": data_doc.strftime("%d/%m/%Y"),
        "data_documento": data_doc.strftime("%d/%m/%Y"),
        "header_titolo": "Relazione Tecnica - Impianto Elettrico (DiCo)",
        "progettista_blocco": progettista_blocco,
        "progettista_nome": progettista_nome,
        "premessa": premessa,
        "norme": norme,
        "criterio_progetto": criterio_testo,
        "dati_tecnici": dati_tecnici,
        "descrizione_impianto": descrizione_impianto,
        "confini": confini_txt,
        "quadri": quadri_list,
        "linee": linee_list,
        "sicurezza": sicurezza,
        "verifiche": verifiche,
        "manutenzione": manutenzione,
        "allegati": allegati,
        "disclaimer_calcoli": "Calcoli e verifiche riportati sono di sintesi e a supporto documentale. Non sostituiscono un progetto esecutivo completo né le verifiche previste dalle norme applicabili.",
        "titolo_cover": "RELAZIONE TECNICA - IMPIANTO ELETTRICO (DiCo)",
        "sottotitolo_cover": oggetto,
        "nome_progetto": nome_progetto,
        "cover_style": ("engineering" if cover_style.startswith("Engineering") else "legacy"),
        "firma": firmatario,
        "timbro_bytes": timbro_bytes,
        "foto1_bytes": foto1_bytes,
        "foto2_bytes": foto2_bytes,
        "foto3_bytes": foto3_bytes,
        "foto4_bytes": foto4_bytes,
        "oggetto_intervento": oggetto,
        "tipologia": tipologia,
        "sistema": sistema,
        "tensione": tensione,
        "potenza_disp": potenza_disp_kw,
        "cod_progetto": cod_progetto,
        "n_doc": n_doc,
        "n_documento": n_doc,
        "rev": revisione,
        "revisione": revisione,
        "impresa": impresa,
        "luogo_firma": luogo_firma,
        "data_firma": data_firma.strftime("%d/%m/%Y"),
    }

    pdf_bytes = genera_pdf_relazione_bytes(payload)
    st.success("PDF generato.")
    st.download_button(
        "Scarica PDF",
        data=pdf_bytes,
        file_name="Relazione_Tecnica_DiCo_Impianto_Elettrico.pdf",
        mime="application/pdf",
    )
