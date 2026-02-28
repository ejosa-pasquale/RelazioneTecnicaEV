# Relazione tecnica EV — Streamlit (v11)

Questa versione fa due cose in modo robusto:

1) **Crea punti di aggancio nel corpo del testo** (se mancanti) inserendo automaticamente i marker:
- {{DITTA_ESECUTRICE}}
- {{LAYOUT_DESCRITTIVO}}
- {{COLONNINE}}
- {{FOTO_GALLERY}}
- {{DIAGRAMMA_IMPIANTO}}
- {{ALLEGATI_SCHEDA_TECNICA}}

2) Genera il DOCX popolando **il corpo**:
- sostituisce placeholder base (es. {{OGGETTO}}, {{COMMITTENTE}}, …) se presenti
- scrive le sezioni (layout/colonnine/foto/diagramma/allegati) nei punti di aggancio
- pulisce la sezione Layout (cancella il testo “non pertinente” sotto il titolo e reinserisce Inclusi/Esclusi)

## Uso rapido
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Suggerimento operativo
- Carica il tuo template DOCX.
- Clicca **“Prepara template (agganci)”**.
- Scarica e salva il “template_preparato.docx”: sarà quello che userai sempre.
