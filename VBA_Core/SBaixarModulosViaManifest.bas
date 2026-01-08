Attribute VB_Name = "SBaixarModulosViaManifest"
Private Sub BaixarModulosViaManifest()
    Dim urlManifest As String
    Dim baseUrl As String
    Dim pastaTemp As String
    Dim linhas As Variant
    Dim i As Long
    Dim conteudo As String
    Dim nomeArquivo As String

    ' Debug.Print "Dentro BaixarModulosViaManifest"

    baseUrl = "https://raw.githubusercontent.com/leveragesecuritizadora/VBA_functions/emissao_unica/"
    urlManifest = baseUrl & "manifest.txt"
    pastaTemp = Environ("TEMP") & "\vba\"

    If Dir(pastaTemp, vbDirectory) = "" Then MkDir pastaTemp

    conteudo = BaixarTexto(urlManifest)
    If conteudo = "" Then
        MsgBox "Erro ao baixar manifest.txt", vbCritical
        Exit Sub
    End If

    linhas = Split(conteudo, vbLf)
    ' Debug.Print "Meio 1 BaixarModulosViaManifest"

    ' Limpando pasta
    On Error Resume Next
    Kill pastaTemp & "*.bas"
    On Error GoTo 0
    ' Debug.Print "Meio 2 BaixarModulosViaManifest"

    LimparTerminal "Baixando módulos"

    Dim urlArquivo As String
    Dim nomeArquivoFormatado As String
    Dim nTotalArquivos As Integer
    Dim iArquivo As Integer

    nTotalArquivos = UBound(linhas) + 1

    For i = LBound(linhas) To UBound(linhas)
        nomeArquivo = Trim(linhas(i))
        iArquivo = i+1


        urlArquivo = baseUrl & nomeArquivo
        nomeArquivoFormatado = pastaTemp & Replace(nomeArquivo, "VBAs", "")
        nomeArquivoFormatado = pastaTemp & Replace(nomeArquivo, "/", "")

        If nomeArquivo <> "" Then
            If Not BaixarArquivo(urlArquivo, nomeArquivoFormatado) Then
                Debug.Print "Erro ao baixar " & nomeArquivo, vbCritical
            Else
                Debug.Print iArquivo &"/"& nTotalArquivos & " - " &  nomeArquivo
            End If
        End If
    Next i
    ' Debug.Print "Saindo BaixarModulosViaManifest"

End Sub