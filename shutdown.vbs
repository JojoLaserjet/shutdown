Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Pfad zum Bild
imagePath = "C:\Users\jojol\Downloads\shutdown-main\shutdown.png" ' Passe den Pfad entsprechend an

' Nachricht anzeigen, wenn das Bild auf dem Desktop geklickt wird
If objFSO.FileExists(imagePath) Then
    result = objShell.Popup("Du hast jetzt viel gearbeitet! Jetzt ist Feierabend!", 0, "Feierabend", 36)
    If result = 6 Then ' Ja
        objShell.Run "shutdown.exe /s /t 0"
    End If
End If
