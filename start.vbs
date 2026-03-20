Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\SOFTWARE\CDR\ETIKETTEN"
WshShell.Run "cmd /c python app.py", 0, False
WScript.Sleep 1500
WshShell.Run "cmd /c start """" http://localhost:5000", 0, False
