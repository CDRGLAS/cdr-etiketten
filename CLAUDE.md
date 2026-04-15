# ETIKETTEN

CDR Glas AG – Handelsetiketten-App (Flask Web-App).
Druckt Glas-Handelsetiketten auf CAB-Etikettendrucker.

## Architektur
- `app.py` – Flask-Server mit eingebettetem HTML, liest AUF.xlsx und KUNDEN.xlsx von R:\CDR-Glas\Lagerverwaltung
- `etiketten.py` – Standalone Tkinter-GUI (alt)
- `index.html` – Standalone HTML-Version mit Browser-Druck
- Druck via win32com (Excel COM) an CAB-EOS5/200 Drucker

## Start
- `autostart.vbs` – Wird per Scheduled Task bei Anmeldung gestartet (kein Browser)
- `start.vbs` – Manueller Start mit Browser
- App läuft auf Port 5000 (Fallback auf 5001+ wenn belegt)
- Mitarbeiter greifen über http://192.168.67.184:5000 zu
