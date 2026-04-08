# BAP Processor & Dashboard

Sistema di monitoraggio e analisi per la produzione Tanzi (BAP). L'applicazione elabora dati da file SAP ed Excel per generare una dashboard interattiva di controllo produzione e inventari.

## Funzionalità

- **Elaborazione Dati**: Trasforma report SAP ZPP093 e matrici P-NUM in dati strutturati.
- **Dashboard Interattiva**: Visualizzazione semaforica dei componenti, WIP e quantità finite.
- **Gestione Archivi**: Possibilità di salvare istantanee della produzione localmente o su Supabase.
- **Override Manuali**: Interfaccia per correggere quantità e impostare target di produzione.

## Requisiti

- Python 3.x
- Dipendenze (installabili via pip):
  ```bash
  pip install pandas openpyxl numpy
  ```

## Utilizzo

Per avviare l'applicazione completa (aggiornamento dati + server):

```bash
bash run_bap.sh
```

Oppure in ambiente Windows:

```bash
run_bap.bat
```

### Componenti Principali

- `bap_processor.py`: Motore di calcolo e trasformazione dati.
- `bap_server.py`: Server HTTP locale (default port 8000) per la dashboard.
- `dashboard.html`: Interfaccia frontend generata dinamicamente.

## Configurazione

Il sistema utilizza diversi file JSON per la persistenza e mappatura:
- `bap_mapping_permanent.json`: Mappatura stazioni e materiali.
- `bap_inventory_baseline.json`: Inventario di partenza.
- `bap_config.json`: (Opzionale) Configurazioni per l'archiviazione su cloud.
