@echo off
REM CDR Etiketten - Watchdog
REM Wird per Scheduled Task alle 5 Minuten gestartet.
REM Prueft per Health-Check ob die App laeuft. Wenn nicht, ruft autostart.vbs auf.

curl.exe -s -o nul -m 3 http://localhost:5000/api/health
if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] Watchdog: App nicht erreichbar - triggere autostart.vbs >> "C:\SOFTWARE\CDR\ETIKETTEN\autostart_log.txt"
    wscript.exe "C:\SOFTWARE\CDR\ETIKETTEN\autostart.vbs"
)
