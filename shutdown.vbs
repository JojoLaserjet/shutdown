Set objShell = CreateObject("WScript.Shell")

result = objShell.Popup("Du hast jetzt viel gearbeitet! Jetzt ist Feierabend!", 0, "Feierabend", 36)
If result = 6 Then 
    objShell.Run "shutdown.exe /s /t 0"
End If
