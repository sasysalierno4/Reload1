# Wind3 Reload Dashboard – A6

Dashboard Streamlit per l'analisi settimanale dell'avanzamento Reload in Campania e Puglia.

## Funzionalità

- Upload file Excel avanzamento
- Tabella riepilogativa per District Manager con KPI e trend
- Soglie critiche configurabili dalla sidebar (Gross, Net, Forever)
- Fasce di attenzione personalizzate (es. 60-70%, 70-80%)
- Download messaggi WhatsApp pronti per ogni DM
- Download Excel individuale per ogni DM
- Download ZIP con tutti gli Excel
- Download riepilogo generale

## Deploy su Streamlit Cloud (gratis)

### 1. Crea repository GitHub
1. Vai su https://github.com/new
2. Crea un repo privato (es. `wind3-reload-dashboard`)
3. Carica i file: `app.py` e `requirements.txt`

### 2. Deploy su Streamlit Cloud
1. Vai su https://share.streamlit.io
2. Accedi con GitHub
3. Clicca **New app**
4. Seleziona il repo, branch `main`, file `app.py`
5. Clicca **Deploy**

In 2-3 minuti la webapp è online con un link tipo:
`https://tuonome-wind3-reload.streamlit.app`

### 3. Uso quotidiano
1. Apri la webapp da browser (funziona su Android)
2. Imposta le soglie dalla sidebar sinistra
3. Carica il file Excel
4. Analizza, scarica messaggi e Excel

## Uso in locale

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Configurazione soglie (sidebar)

| Parametro | Default | Descrizione |
|---|---|---|
| Gross % minima | 40% | Sotto → 🔴 critico |
| Net % minima | 37% | Sotto → 🔴 critico |
| Forever = 0 | Sì | Zero store forever → critico |
| Fasce personalizzate | Off | Range da attenzionare (es. 60-70%) |
| Ambassador code | A6 | Codice zona |
| Regioni | Campania, Puglia | Filtro geografico |
