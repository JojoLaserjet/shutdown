Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

desktopPath = objShell.SpecialFolders("Desktop")
imagePath = desktopPath & "C:\Users\admin\Downloads\shutdown-main\shutdown.png" '

If objFSO.FileExists(imagePath) Then
    result = objShell.Popup("Du hast jetzt viel gearbeitet! Jetzt ist Feierabend!", 0, "Feierabend", 4 + 32)
    If result = 6 Then ' Ja
        objShell.Run "shutdown.exe /s /t 0"
    End If
End If
