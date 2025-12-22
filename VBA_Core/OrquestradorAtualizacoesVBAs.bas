Attribute VB_Name = "OrquestradorAtualizacoesVBAs"

Public Sub OrquestradorAtualizacoesVBAs()
    On Error GoTo Fim
    
    Application.ScreenUpdating = False

    Call ApagarModulos
    Call BaixarModulosViaManifest
    Call ImportarModulos

    Application.ScreenUpdating = True
    MsgBox "Projeto VBA atualizado com sucesso!", vbInformation

    Fim:
        Application.ScreenUpdating = True
End Sub

Private Function BaixarArquivoGitHubPublico(url As String, caminhoLocal As String) As Boolean
    Dim http As Object
    Dim stream As Object

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send

    If http.Status <> 200 Then
        BaixarArquivoGitHubPublico = False
        Exit Function
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' binário
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile caminhoLocal, 2
    stream.Close

    BaixarArquivoGitHubPublico = True
End Function

Private Function BaixarTexto(url As String) As String
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

Private Sub ApagarModulos()
    Dim i As Long
    Dim vbComp As Object

    For i = ThisWorkbook.VBProject.VBComponents.Count To 1 Step -1
        Set vbComp = ThisWorkbook.VBProject.VBComponents(i)

        If vbComp.Type = 1 _
           And vbComp.Name <> "Bootloader" _
           And vbComp.Name <> "OrquestradorAtualizacoesVBAs" Then
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
        End If
    Next i
End Sub

Private Sub BaixarModulosViaManifest()
    Dim urlManifest As String
    Dim baseUrl As String
    Dim pastaTemp As String
    Dim linhas As Variant
    Dim i As Long
    Dim conteudo As String
    Dim nomeArquivo As String

    baseUrl = "https://raw.githubusercontent.com/leveragesecuritizadora/VBA_functions/main/"
    urlManifest = baseUrl & "manifest.txt"
    pastaTemp = Environ("TEMP") & "\vba\"

    If Dir(pastaTemp, vbDirectory) = "" Then MkDir pastaTemp

    conteudo = BaixarTexto(urlManifest)
    If conteudo = "" Then
        MsgBox "Erro ao baixar manifest.txt", vbCritical
        Exit Sub
    End If

    linhas = Split(conteudo, vbLf)

    ' Limpando pasta
    On Error Resume Next
    Kill pastaTemp & "*.bas"
    On Error GoTo 0


    For i = LBound(linhas) To UBound(linhas)
        nomeArquivo = Trim(linhas(i))

        If nomeArquivo <> "" Then
            If Not BaixarArquivoGitHubPublico(baseUrl & nomeArquivo, pastaTemp & nomeArquivo) Then
                MsgBox "Erro ao baixar " & nomeArquivo, vbCritical
            End If
        End If
    Next i
End Sub

Private Sub ImportarModulos()
    Dim pasta As String
    Dim arquivo As String

    pasta = Environ("TEMP") & "\vba\"
    arquivo = Dir(pasta & "*.bas")

    Do While arquivo <> ""
        If arquivo <> "OrquestradorAtualizacoesVBAs.bas" Then
            ThisWorkbook.VBProject.VBComponents.Import pasta & arquivo
        End If
        arquivo = Dir
    Loop
End Sub