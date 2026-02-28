\
# Generatore Relazione Progetto Elettrico (Streamlit)

Web app Streamlit per inserire i dati di progetto e generare una relazione tecnica **in formato DOCX** a partire da un template.

> Il template incluso è basato sul DOCX che hai fornito. La generazione avviene tramite sostituzioni testuali nelle parti standard del documento.
> Nota: eventuali contenuti inseriti in **textbox/forme Word** potrebbero non essere modificabili con `python-docx` e quindi richiedere un template senza textbox oppure un adattamento.

## Avvio in locale

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

Apri il browser sull'URL mostrato da Streamlit.

## Template

- Template di default: `templates/relazione_base.docx`
- Puoi caricare un template alternativo dall'app (sidebar), che sovrascrive il file locale.

## Deploy su Streamlit Community Cloud

1. Crea un repo su GitHub e carica questi file.
2. Vai su Streamlit Community Cloud → **New app**.
3. Seleziona repo e branch.
4. **Main file path**: `app.py`
5. Deploy.

## Struttura progetto

```
.
├── app.py
├── generator.py
├── requirements.txt
└── templates
    └── relazione_base.docx
```
