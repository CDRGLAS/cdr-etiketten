@echo off
REM CDR Etiketten - Auto-Restart-Wrapper
REM Startet python app.py in einer Schleife. Nach 5 schnellen Crashes (<30s)
REM wird abgebrochen, damit kein heisser Endlos-Loop entsteht.

setlocal enabledelayedexpansion
cd /d "C:\SOFTWARE\CDR\ETIKETTEN"
set FAST_CRASHES=0

:loop
for /f "tokens=1-3 delims=:.," %%a in ("!TIME!") do set /a SEC_START=%%a*3600+%%b*60+%%c
echo [%date% %time%] Starte python app.py >> autostart_log.txt

C:\Python314\python.exe app.py 2>> crash_log.txt
set EXIT=%ERRORLEVEL%

for /f "tokens=1-3 delims=:.," %%a in ("!TIME!") do set /a SEC_END=%%a*3600+%%b*60+%%c
set /a DURATION=SEC_END-SEC_START
if !DURATION! LSS 0 set /a DURATION+=86400

echo [%date% %time%] python beendet exit=!EXIT! nach !DURATION!s >> autostart_log.txt

if !DURATION! LSS 30 (
    set /a FAST_CRASHES+=1
) else (
    set FAST_CRASHES=0
)

if !FAST_CRASHES! GEQ 5 (
    echo [%date% %time%] FEHLER: 5 schnelle Crashes - Wrapper bricht ab >> autostart_log.txt
    exit /b 1
)

REM Kurze Pause vor Restart
ping -n 3 127.0.0.1 >nul
goto loop
