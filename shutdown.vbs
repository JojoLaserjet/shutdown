Set objShell = CreateObject("WScript.Shell")

' Nachricht anzeigen
result = objShell.Popup("Du hast jetzt viel gearbeitet! Jetzt ist Feierabend!", 0, "Feierabend", 36)
If result = 6 Then ' Ja
    objShell.Run "shutdown.exe /s /t 0"
End If
