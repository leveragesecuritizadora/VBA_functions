Attribute VB_Name = "SBaixarTexto"

Function BaixarTexto(url As String) As String
    Dim http As Object

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send

    If http.Status = 200 Then
        BaixarTexto = http.responseText
    Else
        BaixarTexto = ""
    End If
End Function