import streamlit as st
import pandas as pd
from datetime import date

from calcoli import corrente_da_potenza, caduta_tensione, verifica_tt_ra_idn, zs_massima_tn
from pdf_generator import genera_pdf_relazione_bytes

st.set_page_config(page_title="Relazione Tecnica DiCo – Impianti Elettrici", layout="wide")

# =========================
# Header
# =========================
st.title("Relazione Tecnica per DiCo – Impianto Elettrico")
st.caption("Interfaccia unica: compilazione + verifiche essenziali + generazione PDF (stile v7).")

PROGETTISTA_BLOCCO = """Ing. Pasquale Senese
via Francesco Soave 30 – 20151 Milano
Cell. 340 5731381 – Email: pasquale.senese@ingpec.eu
P.IVA 14572980960
"""

# =========================
# Sidebar
# =========================
with st.sidebar:
    st.header("Output")
    dv_lim_default = 4.0
    dv_lim = st.number_input("Limite caduta di tensione (%)", min_value=1.0, max_value=10.0, value=dv_lim_default, step=0.5)
    ul_tt = st.number_input("UL TT (V) – criterio Ra·Idn ≤ UL", min_value=25.0, max_value=100.0, value=50.0, step=5.0)
    st.divider()
    st.markdown("**Nota**: le verifiche sono di sintesi (supporto alla DiCo). In casi complessi usare progetto di calcolo dedicato.")

# =========================
# 1) Dati anagrafici
# =========================
st.subheader("1) Dati generali")
c1, c2 = st.columns([2,2])
with c1:
    committente = st.text_input("Committente (Nome/Cognome o Ragione sociale)", "XXXX (Inserire)")
    impianto_indirizzo = st.text_input("Indirizzo impianto", "XXXX (Inserire)")
    data_doc = st.date_input("Data", value=date.today())
with c2:
    impresa = st.text_input("Impresa installatrice", "XXXX (Inserire)")
    impresa_rea = st.text_input("REA / CCIAA (se disponibile)", "XXXX (Inserire)")
    dico_rif = st.text_input("Riferimento DiCo (numero/data)", "XXXX (Inserire)")

# =========================
# 2) Dati tecnici minimi
# =========================
st.subheader("2) Dati tecnici essenziali (minimi)")
t1, t2, t3, t4 = st.columns(4)
with t1:
    alimentazione = st.selectbox("Alimentazione", ["Monofase 230 V", "Trifase 400 V"], index=1)
with t2:
    sistema = st.selectbox("Sistema di distribuzione", ["TT", "TN-S", "TN-C-S"], index=0)
with t3:
    potenza_tot_kw = st.number_input("Potenza prevista/servita (kW)", min_value=0.5, max_value=500.0, value=6.0, step=0.5)
with t4:
    cosphi = st.number_input("cosφ (se noto)", min_value=0.3, max_value=1.0, value=0.95, step=0.01)

Ib = corrente_da_potenza(potenza_tot_kw, alimentazione, cosphi=cosphi)

st.info(f"Corrente di impiego indicativa Ib ≈ **{Ib:.1f} A** (da potenza inserita).")

# =========================
# 3) Ambienti e confini
# =========================
st.subheader("3) Destinazione d’uso / ambienti e confini intervento")
a1, a2 = st.columns(2)
with a1:
    ambienti = st.multiselect(
        "Seleziona ambienti/condizioni (checklist)",
        ["Ordinario", "Bagno", "Esterno", "Locale tecnico", "Autorimessa", "Cantiere", "Altro"],
        default=["Ordinario"],
    )
    ambiente_altro = ""
    if "Altro" in ambienti:
        ambiente_altro = st.text_input("Specificare 'Altro'", "XXXX (Inserire)")
with a2:
    compresi = st.text_area("Intervento comprende", "XXXX (Inserire)", height=110)
    esclusi = st.text_area("Intervento esclude", "XXXX (Inserire)", height=110)
    integrazione = st.selectbox("Integrazione con impianto esistente", ["No", "Sì"], index=1)

# =========================
# 4) Quadri
# =========================
st.subheader("4) Quadri (sintesi)")
default_quadri = pd.DataFrame([
    {"Quadro":"QG", "Ubicazione":"XXXX (Inserire)", "IP":"XX", "Generale":"MT In=XXA curva X Icn XXkA", "Diff":"Tipo A 30mA (se presente)"},
])
quadri_df = st.data_editor(default_quadri, num_rows="dynamic", use_container_width=True, key="quadri")

# =========================
# 5) Linee/circuiti + calcoli
# =========================
st.subheader("5) Linee / circuiti (con calcolo ΔV e verifiche base)")
st.caption("Compila in modo leggero: Cavo + Protezione + Differenziale. Il calcolo principale qui è la caduta di tensione.")

default_linee = pd.DataFrame([
    {"Linea":"L1", "Uso":"Prese", "L_m":25, "P_kw":2.0, "Sez_mm2":2.5, "Protezione":"MT 16A curva C", "Diff":"Tipo A 30mA", "Curva":"C", "In_A":16, "Ra_Ohm":30.0, "Idn_A":0.03},
])
linee_df = st.data_editor(default_linee, num_rows="dynamic", use_container_width=True, key="linee")

def valuta_linea(row):
    # Corrente per linea (se P_kw presente)
    p = float(row.get("P_kw") or 0.0)
    l = float(row.get("L_m") or 0.0)
    s = float(row.get("Sez_mm2") or 0.0)
    curva = str(row.get("Curva") or "C")
    InA = float(row.get("In_A") or 0.0)

    ib = corrente_da_potenza(p, alimentazione, cosphi=cosphi) if p > 0 else 0.0
    dv = caduta_tensione(ib if ib>0 else Ib, l, s, alimentazione, cosphi=cosphi)
    esito = "OK" if dv.delta_v_percent <= dv_lim else "ΔV"

    # TT check (se dati)
    note = []
    if sistema == "TT":
        ra = float(row.get("Ra_Ohm") or 0.0)
        idn = float(row.get("Idn_A") or 0.0)
        if ra > 0 and idn > 0:
            ok_tt = verifica_tt_ra_idn(ra, idn, ul=ul_tt)
            note.append("TT OK" if ok_tt else "TT NO")
            if not ok_tt:
                esito = "TT"
    else:
        # TN check Zs max (solo se In/curva noti)
        if InA > 0:
            zs_max = zs_massima_tn(230.0, curva, InA)
            note.append(f"Zs_max≈{zs_max:.2f}Ω")
    return dv.delta_v_percent, esito, "; ".join(note)

out = []
for _, r in linee_df.iterrows():
    dvp, esito, note = valuta_linea(r)
    out.append((dvp, esito, note))

linee_df_calc = linee_df.copy()
linee_df_calc["ΔV_%"] = [round(x[0],2) for x in out]
linee_df_calc["Esito"] = [x[1] for x in out]
linee_df_calc["Note"] = [x[2] for x in out]

st.dataframe(linee_df_calc, use_container_width=True)

# =========================
# 6) Terra / SPD / sicurezza (descrittivo)
# =========================
st.subheader("6) Terra, SPD e sicurezza (descrizione)")
s1, s2 = st.columns(2)
with s1:
    terra_tipo = st.selectbox("Impianto di terra", ["Nuovo", "Esistente verificato", "Esistente non oggetto di intervento"], index=1)
    equipot = st.selectbox("Collegamenti equipotenziali principali", ["Presenti", "Parziali", "Assenti/da verificare"], index=0)
with s2:
    spd = st.selectbox("SPD (valutazione/installazione)", ["Non previsto", "Previsto", "Presente preesistente"], index=0)
    spd_note = st.text_input("Note SPD (se pertinente)", "XXXX (Inserire)")

# =========================
# 7) Verifiche e prove – registro sintetico
# =========================
st.subheader("7) Verifiche e prove (registro sintetico)")
ver_df = pd.DataFrame([
    {"Prova":"Continuità PE", "Esito":"Conforme", "Strumento":"XXXX (Inserire)"},
    {"Prova":"Isolamento", "Esito":"Conforme", "Strumento":"XXXX (Inserire)"},
    {"Prova":"RCD (tempo/corrente)", "Esito":"Conforme", "Strumento":"XXXX (Inserire)"},
    {"Prova":"Terra (Ra) / Loop (Zs)", "Esito":"Conforme", "Strumento":"XXXX (Inserire)"},
    {"Prova":"Polaritá / Sequenza fasi", "Esito":"Conforme", "Strumento":"XXXX (Inserire)"},
])
ver_df = st.data_editor(ver_df, num_rows="dynamic", use_container_width=True, key="verifiche")

# =========================
# 8) Generazione PDF
# =========================
st.subheader("8) Genera PDF Relazione")

premessa = f"""La presente relazione tecnica è redatta a supporto della Dichiarazione di Conformità (DiCo) rilasciata dall’impresa installatrice.
Non sostituisce la DiCo né i relativi allegati obbligatori.
Assunzioni/Fonte dati: potenza disponibile e condizioni di esercizio fornite dal Committente e/o rilevate in sito in data {data_doc.strftime('%d/%m/%Y')}.
Riferimento DiCo: {dico_rif}.
"""

norme = """- Legge 186/68 (regola dell’arte).
- D.M. 37/08 (ove applicabile).
- CEI 64-8: Parte 4-41, 4-43, 5-52, 5-53, 5-54 (e sezioni pertinenti).
- Norme di prodotto CE / EN per componenti e quadri (ove applicabile).
"""

amb_txt = ", ".join([a for a in ambienti if a != "Altro"])
if "Altro" in ambienti:
    amb_txt += f", Altro: {ambiente_altro}"

dati_tecnici = f"""Alimentazione: {alimentazione}
Sistema di distribuzione: {sistema}
Potenza prevista/servita: {potenza_tot_kw:.1f} kW (Ib indicativa ≈ {Ib:.1f} A a cosφ={cosphi:.2f})
Ambienti/condizioni: {amb_txt}
Impresa installatrice: {impresa} (REA/CCIAA: {impresa_rea})
"""

descrizione_impianto = """L’impianto elettrico oggetto di intervento è realizzato/adeguato nel rispetto della regola dell’arte,
con componenti marcati CE e idonei all’ambiente di installazione. Le condutture sono posate con idonee modalità
(canali/tubazioni/passaggi), con identificazione dei circuiti e adeguata protezione meccanica.
I quadri elettrici sono dotati di dispositivi di manovra e protezione e di opportuna etichettatura.
"""

confini = f"""Intervento comprende: {compresi}
Intervento esclude: {esclusi}
Integrazione con impianto esistente: {integrazione}
"""

sicurezza = f"""Protezione contro i contatti diretti: mediante isolamento delle parti attive e involucri/barriere con grado di protezione adeguato.
Protezione contro i contatti indiretti: mediante interruzione automatica dell’alimentazione (RCD e/o dispositivi di protezione),
collegamento a terra delle masse e collegamenti equipotenziali.
Impianto di terra: {terra_tipo}; equipotenzialità: {equipot}.
SPD: {spd}. Note: {spd_note}.
Criteri di sovracorrente: protezioni magnetotermiche dimensionate in coerenza con le linee (Ib/Sezione) e con potere di interruzione adeguato.
Caduta di tensione: verificata in forma sintetica; limite adottato: {dv_lim:.1f}%.
"""

# Verifiche testo
ver_rows = []
for _, r in ver_df.iterrows():
    ver_rows.append(f"- {r.get('Prova','')}: {r.get('Esito','')} (Strumento: {r.get('Strumento','')})")
verifiche = "Le verifiche finali sono eseguite secondo CEI 64-8 Parte 6 (ove applicabile) e includono, in sintesi:\n" + "\n".join(ver_rows)

manutenzione = """Si raccomanda l’esecuzione di controlli periodici e la verifica dell’efficienza dei dispositivi differenziali (tasto TEST),
oltre alla conservazione della documentazione tecnica e degli schemi aggiornati.
"""

# Allegati
allegati = """Obbligatori (DiCo):
- Schema unifilare / schema dei quadri (obbligatorio).
- Planimetria con ubicazione principali componenti e linee (obbligatorio).
- Elenco materiali principali / dichiarazioni conformità componenti (se previsto).
Consigliati:
- Report fotografico essenziale (quadri, etichette, collegamenti di terra).
- Rapporti prova strumenti (se disponibili).
"""

if st.button("Genera PDF"):
    # build tabelle per pdf
    quadri_list = quadri_df.to_dict(orient="records")
    linee_list = []
    for _, r in linee_df_calc.iterrows():
        linee_list.append({
            "Linea": r.get("Linea",""),
            "Uso": r.get("Uso",""),
            "L_m": r.get("L_m",""),
            "Cavo": f"Cu {r.get('Sez_mm2','')} mm²",
            "Protezione": r.get("Protezione",""),
            "Diff": r.get("Diff",""),
            "DV_perc": f"{r.get('ΔV_%','')}",
            "Esito": r.get("Esito",""),
        })

    payload = {
        "committente_nome": committente,
        "impianto_indirizzo": impianto_indirizzo,
        "data": data_doc.strftime("%d/%m/%Y"),
        "progettista_blocco": PROGETTISTA_BLOCCO,
        "premessa": premessa,
        "norme": norme,
        "dati_tecnici": dati_tecnici,
        "descrizione_impianto": descrizione_impianto,
        "confini": confini,
        "quadri": quadri_list,
        "linee": linee_list,
        "sicurezza": sicurezza,
        "verifiche": verifiche,
        "manutenzione": manutenzione,
        "allegati": allegati,
        "firma": "Ing. Pasquale Senese",
    }

    pdf_bytes = genera_pdf_relazione_bytes(payload)
    st.success("PDF generato.")
    st.download_button(
        "Scarica PDF",
        data=pdf_bytes,
        file_name="Relazione_Tecnica_DiCo_Impianto_Elettrico.pdf",
        mime="application/pdf",
    )
