On Error Resume Next

Do
    Err.Clear

    ' Check for stop file
    stopFile = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\stop_cat_wallpaper.txt"
    If CreateObject("Scripting.FileSystemObject").FileExists(stopFile) Then
        WScript.Echo "Stop file detected. Exiting loop."
        Exit Do
    End If

    ' Fetch JSON from The Cat API
	Set http = CreateObject("MSXML2.XMLHTTP")
	http.Open "GET", "https://api.thecatapi.com/v1/images/search", False
	http.setRequestHeader "x-api-key", "live_UVrCOx3MMUGASRzy1gI7jRDiniAare8XgarVHZ2HgPZ44QDSSHSmcOvc1WnDUw3C"
	http.Send


    If Err.Number = 0 And http.Status = 200 Then
        ' Extract image URL using RegExp
        Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = """url"":""([^""]+)"""
        regex.Global = False
        regex.IgnoreCase = True

        If regex.Test(http.responseText) Then
            Set matches = regex.Execute(http.responseText)
            imageUrl = matches(0).SubMatches(0)

            ' Download the image
            Set http2 = CreateObject("MSXML2.XMLHTTP")
            http2.Open "GET", imageUrl, False
            http2.Send

            If Err.Number = 0 And http2.Status = 200 Then
                Set stream = CreateObject("ADODB.Stream")
                stream.Type = 1 ' Binary
                stream.Open
                stream.Write http2.responseBody

                imagePath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\cat_wallpaper.jpg"
                stream.SaveToFile imagePath, 2
                stream.Close

                ' Set wallpaper
                Set WshShell = CreateObject("WScript.Shell")
                WshShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", imagePath
                Set oShell = CreateObject("Shell.Application")
                oShell.ShellExecute "RUNDLL32.EXE", "USER32.DLL,UpdatePerUserSystemParameters", "", "open", 1
            End If
        End If
    End If

    ' Wait 6.7 seconds before next attempt
    WScript.Sleep 6700
Loop

MsgBox "Cat wallpaper loop stopped.", vbInformation
