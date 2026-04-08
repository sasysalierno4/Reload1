# Wind3 Reload Dashboard — v4

Dashboard Streamlit per il monitoraggio dell'iniziativa **Reload** sulla zona Ambassador.

## 🆕 Novità v4

### Bug fix
- **Filtro zona AND** (non più OR): Ambassador *e* Regione, non *o*. Selezionabile dalla sidebar.
- **`latest` calcolato dopo il filtro**: prima sceglieva il mese più recente sull'intero file, anche se non c'erano dati per la tua zona in quel mese.
- **Divisione per zero protetta** ovunque (`safe_pct`): niente più `inf`/`NaN` sui DM con TAM=0.
- **Soglie davvero usate** in Excel e PDF (prima erano hard-coded a 40/60 e 37/50).
- **Forever = 0 → critico** ora rispetta davvero il flag della sidebar.
- **PDF unicode-safe**: niente più crash su `–`, `→`, `▲`, lettere accentate.
- **Check colonne** all'apertura del file: errore esplicito se manca qualcosa, invece di crash criptico.
- **format_func DM ottimizzato** (precomputed dict, era O(n²) sui selectbox).
- **Cache più solida**: i KPI sono ricalcolati fuori dalla cache, le soglie cambiano subito senza re-upload.

### Nuove funzionalità
- **Sidebar parametrica completa**:
  - Ambassador code (A1, A2, A6, ...)
  - Regioni (lista libera)
  - Modalità filtro: `AND` / `OR` / solo Ambassador / solo Regione
  - Soglia *critica* Gross/Net
  - Soglia *attenzione* Gross/Net (livello intermedio)
  - Flag Forever=0 critico
  - Nome foglio Excel
- **Filtri Tab Riepilogo**: Status, Regione, Tipo Store, Area Manager
- **Ricerca store potenziata**: per nome, ragione sociale, città, indirizzo, provincia + filtro tipo store + filtro regione + colonna AM nei risultati
- **🌳 Tab Mappa Zona** (nuovo): vista ad albero gerarchica AM → DM → Store con KPI per ogni livello, filtri AM e tipo store, esportazione CSV completa.

## 📁 Struttura

```
.
├── app.py                  # Applicazione principale
├── requirements.txt        # Dipendenze Python
├── runtime.txt             # Pinning Python 3.11 per Streamlit Cloud
├── .streamlit/
│   └── config.toml         # Tema e config server
└── README.md
```

## 🚀 Deploy locale

```bash
pip install -r requirements.txt
streamlit run app.py
```

## ☁️ Deploy Streamlit Cloud

1. Push su GitHub (`sasysalierno4/Reload1`).
2. Il file `runtime.txt` forza Python 3.11 (risolve il problema 3.14.3 vs 3.11).
3. Il file `.streamlit/config.toml` applica il tema dark Wind3.
4. Streamlit Cloud rileva `app.py` e parte automaticamente.

## 📊 Formato file Excel atteso

Foglio: `Sales x Store` (configurabile dalla sidebar)

Colonne obbligatorie:
- `SHOP_CODE`, `STORE`, `COMPANY_NAME`, `CITY`, `PROVINCE_CODE`, `STORE_TYPE`
- `REGION`, `AMBASSADOR`, `AREA_MANAGER`, `DISTRICT_MANAGER`
- `TAM`, `SR_PLUS_GROSS_SALES`, `SR_PLUS_NET_SALES`
- `Activeforever`, `TotalStore`, `MONTH`, `YEAR`

Colonne opzionali:
- `STORE_ADDRESS` (usata nella ricerca, creata vuota se assente)
