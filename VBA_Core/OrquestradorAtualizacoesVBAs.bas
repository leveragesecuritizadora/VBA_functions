Attribute VB_Name = "OrquestradorAtualizacoesVBAs"

Public Sub OrquestradorAtualizacoesVBAs()
    Debug.Print "1059Rodando orquestrador"
    On Error GoTo Fim

    Application.ScreenUpdating = False

    Call ApagarModulos
    Debug.Print "rodou ApagarModulos"
    Call BaixarModulosViaManifest
    Debug.Print "rodou BaixarModulosViaManifest"
    Call ImportarModulos
    Debug.Print "rodou ImportarModulos"

    Application.ScreenUpdating = True
    MsgBox "Projeto VBA atualizado com sucesso!", vbInformation

    Fim:
        Application.ScreenUpdating = True
End Sub

Function BaixarArquivoGitHubPublico(url As String, caminhoLocal As String) As Boolean
    Dim http As Object
    Dim stream As Object

    Debug.Print "Dentro BaixarArquivoGitHubPublico: "; url

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.SetTimeouts 5000, 5000, 5000, 10000
    http.Open "GET", url, False
    http.Send

    If http.Status <> 200 Then
        Debug.Print "Erro HTTP " & http.Status & ": " & url
        BaixarArquivoGitHubPublico = False
        Exit Function
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile caminhoLocal, 2
    stream.Close

    Debug.Print "Saindo BaixarArquivoGitHubPublico"
    BaixarArquivoGitHubPublico = True
End Function

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

Sub ApagarModulos()
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

Sub BaixarModulosViaManifest()
    Dim urlManifest As String
    Dim baseUrl As String
    Dim pastaTemp As String
    Dim linhas As Variant
    Dim i As Long
    Dim conteudo As String
    Dim nomeArquivo As String

    Debug.Print "Dentro BaixarModulosViaManifest"

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
    Debug.Print "Meio 1 BaixarModulosViaManifest"

    ' Limpando pasta
    On Error Resume Next
    Kill pastaTemp & "*.bas"
    On Error GoTo 0
    Debug.Print "Meio 2 BaixarModulosViaManifest"

    Dim urlArquivo As String
    Dim nomeArquivoFormatado As String

    For i = LBound(linhas) To UBound(linhas)
        nomeArquivo = Trim(linhas(i))

        urlArquivo = baseUrl & nomeArquivo
        nomeArquivoFormatado = pastaTemp & Replace(nomeArquivo, "VBAs", "")
        nomeArquivoFormatado = pastaTemp & Replace(nomeArquivo, "/", "")

        If nomeArquivo <> "" Then
            If Not BaixarArquivo(urlArquivo, nomeArquivoFormatado) Then
                MsgBox "Erro ao baixar " & nomeArquivo, vbCritical
            End If
        End If
    Next i
    Debug.Print "Saindo BaixarModulosViaManifest"

End Sub

Sub ImportarModulos()
    Dim pasta As String
    Dim arquivo As String

    Debug.Print "Dentro ImportarModulos"

    pasta = Environ("TEMP") & "\vba\"
    arquivo = Dir(pasta & "*.bas")
    Debug.Print "meio ImportarModulos"

    Do While arquivo <> ""
        If arquivo <> "OrquestradorAtualizacoesVBAs.bas" Then
            ThisWorkbook.VBProject.VBComponents.Import pasta & arquivo
        End If
        arquivo = Dir
    Loop
    Debug.Print "saindo ImportarModulos"

End Sub

Private Function BaixarArquivo(url As String, destino As String) As Boolean
    Dim http As Object
    Dim stream As Object

    Debug.Print "Dentro BaixarArquivo: "; url

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send

    ' If http.Status <> 200 Then Exit Function
    If http.Status <> 200 Then
        MsgBox "HTTP Status: " & http.Status & vbCrLf & url, vbCritical
        BaixarArquivo = False
        Exit Function
    End If
    Debug.Print "Dentro 2.0 BaixarArquivo: "; url


    Set stream = CreateObject("ADODB.Stream")
        Debug.Print "Dentro 2.1 BaixarArquivo: "; url

    stream.Type = 1
        Debug.Print "Dentro 2.2 BaixarArquivo: "; url

    stream.Open
        Debug.Print "Dentro 2.3 BaixarArquivo: "; url

    stream.Write http.responseBody
        Debug.Print "Dentro 2.4 BaixarArquivo: "; url

    stream.SaveToFile destino, 2
        Debug.Print "Dentro 2.5 BaixarArquivo: "; url

    stream.Close
        Debug.Print "Dentro 2.6 BaixarArquivo: "; url


    Debug.Print "Dentro 3 BaixarArquivo: "; url
    BaixarArquivo = True
End Function