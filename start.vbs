Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\SOFTWARE\CDR\ETIKETTEN"
' Alte Instanzen auf Port 5000 beenden
WshShell.Run "powershell -NoProfile -Command ""Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }""", 0, True
WScript.Sleep 500
' Drucker-Port Ne04: sicherstellen (IP 192.168.67.71)
WshShell.Run "powershell -NoProfile -Command ""if (-not (Get-PrinterPort -Name 'Ne04:' -ErrorAction SilentlyContinue)) { Add-PrinterPort -Name 'Ne04:' -PrinterHostAddress '192.168.67.71' }; Set-Printer -Name 'CAB-EOS5/200' -PortName 'Ne04:'""", 0, True
WScript.Sleep 500
WshShell.Run "cmd /c python app.py", 0, False
WScript.Sleep 1500
WshShell.Run "cmd /c start """" http://localhost:5000", 0, False
