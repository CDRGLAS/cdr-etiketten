' CDR Etiketten - Watchdog
' Wird per Scheduled Task "CDR Etiketten Watchdog" alle 5 Min ausgefuehrt.
' Prueft per Health-Check ob die App laeuft. Wenn nicht, ruft autostart.vbs auf.
' wscript.exe zeigt kein Fenster (im Gegensatz zu cmd.exe).

Dim http, fso, WshShell

Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

If Not IsAppRunning("http://localhost:5000/api/health") Then
    LogLine "Watchdog: App nicht erreichbar - triggere autostart.vbs"
    WshShell.Run "wscript.exe ""C:\SOFTWARE\CDR\ETIKETTEN\autostart.vbs""", 0, True
End If

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
