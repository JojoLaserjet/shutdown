Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Pfad zum Bild auf dem Desktop
desktopPath = objShell.SpecialFolders("Desktop")
imagePath = desktopPath & "\feierabend.png" ' Hier "feierabend.png" durch den Namen deines Bildes ersetzen

' Nachricht anzeigen, wenn das Bild auf dem Desktop geklickt wird
If objFSO.FileExists(imagePath) Then
    result = objShell.Popup("Du hast jetzt viel gearbeitet! Jetzt ist Feierabend!", 0, "Feierabend", 4 + 32)
    If result = 6 Then ' Ja
        objShell.Run "shutdown.exe /s /t 0"
    End If
End If
