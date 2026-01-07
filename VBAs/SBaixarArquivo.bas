Attribute VB_Name = "SBaixarArquivo"


Function BaixarArquivo(url As String, destino As String) As Boolean
    Dim http As Object
    Dim stream As Object

    ' Debug.Print "Dentro BaixarArquivo: "; url

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send

    ' If http.Status <> 200 Then Exit Function
    If http.Status <> 200 Then
        MsgBox "HTTP Status: " & http.Status & vbCrLf & url, vbCritical
        BaixarArquivo = False
        Exit Function
    End If
    ' Debug.Print "Dentro 2.0 BaixarArquivo: "; url


    Set stream = CreateObject("ADODB.Stream")
        ' Debug.Print "Dentro 2.1 BaixarArquivo: "; url

    stream.Type = 1
        ' Debug.Print "Dentro 2.2 BaixarArquivo: "; url

    stream.Open
        ' Debug.Print "Dentro 2.3 BaixarArquivo: "; url

    stream.Write http.responseBody
        ' Debug.Print "Dentro 2.4 BaixarArquivo: "; url

    stream.SaveToFile destino, 2
        ' Debug.Print "Dentro 2.5 BaixarArquivo: "; url

    stream.Close
        ' Debug.Print "Dentro 2.6 BaixarArquivo: "; url


    ' Debug.Print "Dentro 3 BaixarArquivo: "; url
    BaixarArquivo = True
End Function