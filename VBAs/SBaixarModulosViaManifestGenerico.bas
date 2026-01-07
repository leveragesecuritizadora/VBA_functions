Attribute VB_Name = "SBaixarModulosViaManifestGenerico"

Sub BaixarModulosViaManifestGenerico(url_manifest_generico As string)
    Dim urlManifest As String
    Dim baseUrl As String
    Dim pastaTemp As String
    Dim linhas As Variant
    Dim i As Long
    Dim conteudo As String
    Dim nomeArquivo As String

    Debug.Print "Dentro BaixarModulosViaManifestGenerico"

    baseUrl = "https://raw.githubusercontent.com/leveragesecuritizadora/VBA_functions/main/"
    urlManifest = baseUrl & url_manifest_generico
    pastaTemp = Environ("TEMP") & "\vba\"

    If Dir(pastaTemp, vbDirectory) = "" Then MkDir pastaTemp

    conteudo = BaixarTexto(urlManifest)
    If conteudo = "" Then
        MsgBox "Erro ao baixar " & url_manifest_generico, vbCritical
        Exit Sub
    End If

    linhas = Split(conteudo, vbLf)
    Debug.Print "Meio 1 BaixarModulosViaManifestGenerico"

    ' Limpando pasta
    On Error Resume Next
    Kill pastaTemp & "*.bas"
    On Error GoTo 0
    Debug.Print "Meio 2 BaixarModulosViaManifestGenerico"

    Dim urlArquivo As String
    Dim nomeArquivoFormatado As String

    For i = LBound(linhas) To UBound(linhas)
        nomeArquivo = Trim(linhas(i))

        urlArquivo = baseUrl & nomeArquivo
        nomeArquivoFormatado = pastaTemp & Replace(nomeArquivo, "VBAs", "")
        nomeArquivoFormatado = pastaTemp & Replace(nomeArquivo, "/", "")

        If nomeArquivo <> "" Then
            If Not BaixarArquivo(urlArquivo, nomeArquivoFormatado) Then
                Debug.Print "Erro ao baixar " & nomeArquivo, vbCritical
            End If
        End If
    Next i
    Debug.Print "Saindo BaixarModulosViaManifestGenerico"

End Sub