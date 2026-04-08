#!/bin/bash
echo "============================================"
echo " BAP Processor - Aggiornamento Dashboard"
echo "============================================"
echo

# Vai nella cartella dello script
cd "$(dirname "$0")"

# Installa dipendenze
echo "Verifica dipendenze..."
pip install pandas openpyxl numpy --quiet

echo
echo "Elaborazione in corso..."
python3 bap_processor.py

echo
echo "Avvio Server Interattivo..."
echo "--------------------------------------------"
python3 bap_server.py

echo "Fatto!"
