' CDR Etiketten - Autostart (ohne Browser)
' Startet die App auf dem Server. Prueft zuerst ob sie schon laeuft.

Dim WshShell, fso, http, appPort, healthUrl, maxRetries, i

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
WshShell.CurrentDirectory = "C:\SOFTWARE\CDR\ETIKETTEN"

maxRetries = 10

' --- 1. Pruefe ob App bereits laeuft (Port aus Datei oder Standard 5000) ---
appPort = ReadPort()
If appPort <> "" Then
    healthUrl = "http://localhost:" & appPort & "/api/health"
    If IsAppRunning(healthUrl) Then
        WScript.Quit 0
    End If
End If

If appPort <> "5000" Then
    If IsAppRunning("http://localhost:5000/api/health") Then
        WScript.Quit 0
    End If
End If

' --- 2. Alte Zombie-Prozesse auf Port 5000 beenden ---
WshShell.Run "powershell -NoProfile -Command ""Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }""", 0, True
WScript.Sleep 1000

' --- 3. Drucker-Port sicherstellen ---
WshShell.Run "powershell -NoProfile -Command ""if (-not (Get-PrinterPort -Name 'Ne04:' -ErrorAction SilentlyContinue)) { Add-PrinterPort -Name 'Ne04:' -PrinterHostAddress '192.168.67.71' }; Set-Printer -Name 'CAB-EOS5/200' -PortName 'Ne04:'""", 0, True
WScript.Sleep 500

' --- 4. App starten ---
WshShell.Run "cmd /c C:\Python314\python.exe app.py", 0, False

' --- 5. Warten bis App bereit ist ---
For i = 1 To maxRetries
    WScript.Sleep 1000
    appPort = ReadPort()
    If appPort <> "" Then
        healthUrl = "http://localhost:" & appPort & "/api/health"
        If IsAppRunning(healthUrl) Then
            WScript.Quit 0
        End If
    End If
Next

' --- 6. Fehlermeldung ---
MsgBox "Die Etiketten-App konnte nicht gestartet werden." & vbCrLf & vbCrLf & _
       "Moegliche Ursachen:" & vbCrLf & _
       "  - Port 5000 ist noch von einem anderen Prozess belegt" & vbCrLf & _
       "  - Das Netzlaufwerk R: ist nicht verbunden" & vbCrLf & _
       "  - Python ist nicht installiert" & vbCrLf & vbCrLf & _
       "Details siehe: C:\SOFTWARE\CDR\ETIKETTEN\app_log.txt", _
       vbExclamation, "CDR Etiketten - Startfehler"


Function IsAppRunning(url)
    On Error Resume Next
    http.Open "GET", url, False
    http.setTimeouts 2000, 2000, 2000, 2000
    http.Send
    If Err.Number = 0 And http.Status = 200 Then
        IsAppRunning = True
    Else
        IsAppRunning = False
    End If
    Err.Clear
    On Error GoTo 0
End Function

Function ReadPort()
    Dim portFile, f
    portFile = "C:\SOFTWARE\CDR\ETIKETTEN\app_port.txt"
    ReadPort = ""
    If fso.FileExists(portFile) Then
        On Error Resume Next
        Set f = fso.OpenTextFile(portFile, 1)
        If Err.Number = 0 Then
            ReadPort = Trim(f.ReadAll)
            f.Close
        End If
        Err.Clear
        On Error GoTo 0
    End If
End Function
