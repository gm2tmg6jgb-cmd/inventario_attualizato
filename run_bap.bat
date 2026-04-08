@echo off
echo ============================================
echo  BAP Processor - Aggiornamento Dashboard
echo ============================================
echo.

REM Vai nella cartella dello script
cd /d "%~dp0"

REM Controlla Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRORE: Python non trovato. Installa Python da https://python.org
    pause
    exit /b
)

REM Installa dipendenze se mancano
echo Verifica dipendenze...
pip install pandas openpyxl numpy --quiet

echo.
echo Elaborazione in corso...
python bap_processor.py

echo.
echo Avvio Server Interattivo...
echo --------------------------------------------
python bap_server.py

pause
