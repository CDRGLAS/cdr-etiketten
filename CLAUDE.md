# ETIKETTEN

CDR Glas AG – Handelsetiketten-App (Flask Web-App).
Druckt Glas-Handelsetiketten auf CAB-Etikettendrucker.

## Architektur
- `app.py` – Flask-Server mit eingebettetem HTML, liest AUF.xlsx und KUNDEN.xlsx von R:\CDR-Glas\Lagerverwaltung
- `etiketten.py` – Standalone Tkinter-GUI (alt)
- `index.html` – Standalone HTML-Version mit Browser-Druck
- Druck via win32com (Excel COM) an CAB-EOS5/200 Drucker (IP 192.168.67.71, Port Ne04:)
- Etikette_Vorlage.xlsx wird befüllt und direkt gedruckt

## Netzwerk / Infrastruktur
- App-Server: 192.168.67.184 (Gabriels PC)
- Mitarbeiter (PC_AVOR, 192.168.67.156) greifen per Browser auf http://192.168.67.184:5000 zu
- Datenquelle: R:\CDR-Glas\Lagerverwaltung\AUF.xlsx und KUNDEN.xlsx (Netzlaufwerk)
- Python: C:\Python314\python.exe (für alle Windows-Benutzer verfügbar, NICHT den benutzerspezifischen Python verwenden)

## Start-Mechanismus
- `autostart.vbs` – Wird per Windows Scheduled Task "CDR Etiketten Autostart" bei Anmeldung gestartet (kein Browser)
- `start.vbs` – Manueller Start mit Browser, Desktop-Verknüpfung zeigt hierhin
- Beide Scripts prüfen per Health-Check (`/api/health`) ob die App schon läuft
- App läuft auf Port 5000 (Fallback auf 5001+ wenn belegt)
- Port wird in `app_port.txt` geschrieben
- Logging in `app_log.txt` – bei Problemen immer hier nachschauen

## Bekannte Probleme / Hinweise
- Der Mitarbeiter startet die App remote, aber der Prozess läuft auf Gabriels PC unter dem Benutzer des Mitarbeiters. Dadurch kann Gabriel den Prozess nicht ohne Admin-Rechte beenden (Zugriff verweigert). Die App hat deshalb eine Fallback-Port-Logik.
- Zombie-Prozesse auf Port 5000 waren ein wiederkehrendes Problem. Die VBS-Scripts versuchen alte Prozesse per PowerShell zu beenden, aber das scheitert bei prozessübergreifenden Benutzerrechten.
- VBS-Dateien dürfen keine Umlaute enthalten (ANSI-Encoding, sonst Zeichensalat in MsgBox)
- Dateien (.bat, .vbs, .url) auf dem Netzlaufwerk R: lösen Windows-Sicherheitswarnungen aus. Desktop-Verknüpfungen für Mitarbeiter besser als Browser-Lesezeichen anlegen.

## Sprache
- Etiketten gibt es in DE und FR
- Sprache wird automatisch aus KUNDEN.xlsx ermittelt (Zuordnungsfeld enthält "FR")
- Druck-Log (`druck_log.json`) trackt welche Aufträge/Positionen bereits gedruckt wurden
