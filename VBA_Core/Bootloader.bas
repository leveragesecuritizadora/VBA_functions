Attribute VB_Name = "Bootloader"

Public Sub AtualizarProjeto()
    Dim url As String
    Dim pastaTemp As String
    Dim caminhoCore As String

    url = "https://raw.githubusercontent.com/leveragesecuritizadora/VBA_functions/tree/main/VBA_Core/Core.bas"
    pastaTemp = Environ("TEMP") & "\vba\"
    caminhoCore = pastaTemp & "Core.bas"

    If Dir(pastaTemp, vbDirectory) = "" Then MkDir pastaTemp

    If Not BaixarArquivo(url, caminhoCore) Then
        MsgBox "Falha ao baixar Core.bas", vbCritical
        Exit Sub
    End If

    ImportarCore caminhoCore

    ' chama o orquestrador REAL
    Application.Run "AtualizarProjetoVBA"
End Sub

Private Function BaixarArquivo(url As String, destino As String) As Boolean
    Dim http As Object
    Dim stream As Object

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send

    If http.Status <> 200 Then Exit Function

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile destino, 2
    stream.Close

    BaixarArquivo = True
End Function

Private Sub ImportarCore(caminho As String)
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove _
        ThisWorkbook.VBProject.VBComponents("Core")
    On Error GoTo 0

    ThisWorkbook.VBProject.VBComponents.Import caminho
End Sub
