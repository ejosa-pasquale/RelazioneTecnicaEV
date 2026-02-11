# Relazione Tecnica DiCo – Streamlit App

App Streamlit per compilare una **Relazione Tecnica** (allegato alla DiCo) per impianti elettrici
con campi guidati + verifiche/collaudi essenziali e generazione **PDF**.

## Avvio locale
```bash
python -m venv .venv
source .venv/bin/activate  # (Windows: .venv\Scripts\activate)
pip install -r requirements.txt
streamlit run app.py
```

## Note
- Il PDF è generato direttamente (ReportLab).
- Le verifiche sono volutamente **sintetiche** (supporto alla DiCo) e non sostituiscono un progetto di calcolo completo.
