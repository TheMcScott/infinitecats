Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
stopFile = shell.ExpandEnvironmentStrings("%TEMP%") & "\stop_cat_wallpaper.txt"

' Create the stop file
If Not fso.FileExists(stopFile) Then
    Set file = fso.CreateTextFile(stopFile, True)
    file.WriteLine "Stop triggered at " & Now
    file.Close
End If

' Wait 5 seconds to give the looping script time to detect it
WScript.Sleep 5000

' Delete the stop file
If fso.FileExists(stopFile) Then
    fso.DeleteFile stopFile
End If

MsgBox "Stop signal sent and file deleted.", vbInformation
