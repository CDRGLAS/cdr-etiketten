' CDR Etiketten - Autostart (ohne Browser)
' Startet die App auf dem Server. Prueft zuerst ob sie schon laeuft.

Dim WshShell, fso, http, appPort, healthUrl, maxRetries, i, driveRetries

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
WshShell.CurrentDirectory = "C:\SOFTWARE\CDR\ETIKETTEN"

maxRetries = 10

LogLine "Autostart gestartet"

' --- 0. Initial warten (Login-Zeit, Netzwerk stabilisieren) ---
WScript.Sleep 15000

' --- 1. Pruefe ob App bereits laeuft (Port aus Datei oder Standard 5000) ---
appPort = ReadPort()
If appPort <> "" Then
    healthUrl = "http://localhost:" & appPort & "/api/health"
    If IsAppRunning(healthUrl) Then
        LogLine "App laeuft bereits auf Port " & appPort
        WScript.Quit 0
    End If
End If

If appPort <> "5000" Then
    If IsAppRunning("http://localhost:5000/api/health") Then
        LogLine "App laeuft bereits auf Port 5000"
        WScript.Quit 0
    End If
End If

' --- 2. Warten bis Netzlaufwerk R: verfuegbar ist (bis zu 60s) ---
driveRetries = 0
Do While Not fso.FileExists("R:\CDR-Glas\Lagerverwaltung\AUF.xlsx")
    If driveRetries >= 60 Then
        LogLine "FEHLER: Netzlaufwerk R: nach 60s nicht verfuegbar - Abbruch"
        Exit Do
    End If
    WScript.Sleep 1000
    driveRetries = driveRetries + 1
Loop
If driveRetries > 0 And fso.FileExists("R:\CDR-Glas\Lagerverwaltung\AUF.xlsx") Then
    LogLine "Netzlaufwerk R: nach " & driveRetries & "s verfuegbar"
End If

' --- 3. Alte Zombie-Prozesse auf Port 5000 beenden ---
WshShell.Run "powershell -NoProfile -Command ""Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }""", 0, True
WScript.Sleep 1000

' --- 4. Drucker-Port sicherstellen ---
WshShell.Run "powershell -NoProfile -Command ""if (-not (Get-PrinterPort -Name 'Ne04:' -ErrorAction SilentlyContinue)) { Add-PrinterPort -Name 'Ne04:' -PrinterHostAddress '192.168.67.71' }; Set-Printer -Name 'CAB-EOS5/200' -PortName 'Ne04:'""", 0, True
WScript.Sleep 500

' --- 5. App starten ---
LogLine "Starte python.exe app.py"
WshShell.Run "cmd /c C:\Python314\python.exe app.py", 0, False

' --- 6. Warten bis App bereit ist ---
For i = 1 To maxRetries
    WScript.Sleep 1000
    appPort = ReadPort()
    If appPort <> "" Then
        healthUrl = "http://localhost:" & appPort & "/api/health"
        If IsAppRunning(healthUrl) Then
            LogLine "App bereit auf Port " & appPort & " (nach " & i & "s)"
            WScript.Quit 0
        End If
    End If
Next

' --- 7. Fehlermeldung ---
LogLine "FEHLER: App nach " & maxRetries & "s nicht erreichbar"
MsgBox "Die Etiketten-App konnte nicht gestartet werden." & vbCrLf & vbCrLf & _
       "Moegliche Ursachen:" & vbCrLf & _
       "  - Port 5000 ist noch von einem anderen Prozess belegt" & vbCrLf & _
       "  - Das Netzlaufwerk R: ist nicht verbunden" & vbCrLf & _
       "  - Python ist nicht installiert" & vbCrLf & vbCrLf & _
       "Details siehe: C:\SOFTWARE\CDR\ETIKETTEN\app_log.txt", _
       vbExclamation, "CDR Etiketten - Startfehler"


Sub LogLine(msg)
    Dim logFile, f, ts
    logFile = "C:\SOFTWARE\CDR\ETIKETTEN\autostart_log.txt"
    ts = Year(Now) & "-" & Right("0" & Month(Now),2) & "-" & Right("0" & Day(Now),2) & " " & _
         Right("0" & Hour(Now),2) & ":" & Right("0" & Minute(Now),2) & ":" & Right("0" & Second(Now),2)
    On Error Resume Next
    Set f = fso.OpenTextFile(logFile, 8, True)
    If Err.Number = 0 Then
        f.WriteLine ts & " " & msg
        f.Close
    End If
    Err.Clear
    On Error GoTo 0
End Sub


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
