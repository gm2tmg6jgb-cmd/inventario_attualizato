#!/bin/bash
echo "============================================"
echo " BAP Processor - Aggiornamento Dashboard"
echo "============================================"
echo

# Vai nella cartella dello script
cd "$(dirname "$0")"

# Attivazione ambiente virtuale se presente
if [ -d ".venv" ]; then
    echo "Attivazione ambiente virtuale (.venv)..."
    source .venv/bin/activate
fi

# Installa dipendenze
echo "Verifica dipendenze..."
python3 -m pip install pandas openpyxl numpy --quiet

echo
echo "Elaborazione in corso..."
python3 bap_processor.py

echo
echo "Avvio Server Interattivo..."
echo "--------------------------------------------"
echo "I log verranno salvati in server.log"
# Esegue il server e reindirizza sia stdout che stderr su file E a video
python3 bap_server.py 2>&1 | tee -a server.log

echo "Fatto!"
