Attribute VB_Name = "Bootloader"

Public Sub RodarBootloader()
    Dim url As String
    Dim pastaTemp As String
    Dim caminhoCore As String

    url = "https://raw.githubusercontent.com/leveragesecuritizadora/VBA_functions/main/VBA_Core/SOrquestradorAtualizacoesVBAs.bas"
    pastaTemp = Environ("TEMP") & "\vba\"
    caminhoCore = pastaTemp & "SOrquestradorAtualizacoesVBAs.bas"

    If Dir(pastaTemp, vbDirectory) = "" Then MkDir pastaTemp

    If Not BaixarArquivo(url, caminhoCore) Then
        MsgBox "Falha ao baixar SOrquestradorAtualizacoesVBAs.bas", vbCritical
        Exit Sub
    End If

    ImportarOrquestrador caminhoCore

    ' chama o orquestrador REAL
    Application.Run "OrquestradorAtualizacoesVBAs"
    Debug.Print "aquiiii"
End Sub

Private Function BaixarArquivo(url As String, destino As String) As Boolean
    Dim http As Object
    Dim stream As Object

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send

    ' If http.Status <> 200 Then Exit Function
    If http.Status <> 200 Then
        MsgBox "HTTP Status: " & http.Status & vbCrLf & url, vbCritical
        BaixarArquivo = False
        Exit Function
    End If


    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile destino, 2
    stream.Close

    Debug.Print "Estou aqui"
    BaixarArquivo = True
End Function

Private Sub ImportarOrquestrador(caminho As String)
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove _
        ThisWorkbook.VBProject.VBComponents("OrquestradorAtualizacoesVBAs")
    On Error GoTo 0

    ThisWorkbook.VBProject.VBComponents.Import caminho
End Sub
