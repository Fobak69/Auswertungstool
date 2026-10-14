@echo off
chcp 65001 >nul
title Excel-Pruefer - Korrektur-Assistent
cd /d "%~dp0"

echo ======================================================================
echo          EXCEL-PRUEFER - KORREKTUR-ASSISTENT WIRD GESTARTET
echo ======================================================================
echo.
echo Initialisiere Anwendung...
echo.

python app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [HINWEIS] Versuche Start im Browser-Modus...
    python app.py --browser
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ein Fehler ist aufgetreten. Bitte stellen Sie sicher, dass Python installiert ist.
    pause
)
