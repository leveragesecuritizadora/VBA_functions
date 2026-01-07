Attribute VB_Name = "SBaixarModulosViaManifestGenerico"

Sub BaixarModulosViaManifestGenerico( _ 
    url_manifest_generico As String, 
    pasta_temp As String _
)
    Dim urlManifest As String
    Dim baseUrl As String
    Dim linhas As Variant
    Dim i As Long
    Dim conteudo As String
    Dim nomeArquivo As String

    Debug.Print "Dentro BaixarModulosViaManifestGenerico"

    baseUrl = "https://raw.githubusercontent.com/leveragesecuritizadora/VBA_functions/main/"
    urlManifest = baseUrl & url_manifest_generico

    If Dir(pasta_temp, vbDirectory) = "" Then MkDir pasta_temp

    conteudo = BaixarTexto(urlManifest)
    If conteudo = "" Then
        MsgBox "Erro ao baixar " & url_manifest_generico, vbCritical
        Exit Sub
    End If

    linhas = Split(conteudo, vbLf)
    Debug.Print "Meio 1 BaixarModulosViaManifestGenerico"

    ' Limpando pasta
    On Error Resume Next
    Kill pasta_temp & "*.bas"
    On Error GoTo 0
    Debug.Print "Meio 2 BaixarModulosViaManifestGenerico"

    Dim urlArquivo As String
    Dim nomeArquivoFormatado As String
    Dim nTotalArquivos As Integer
    Dim iArquivo As Integer

    nTotalArquivos = UBound(linhas) + 1

    For i = LBound(linhas) To UBound(linhas)
        nomeArquivo = Trim(linhas(i))
        iArquivo = i+1


        urlArquivo = baseUrl & nomeArquivo
        nomeArquivoFormatado = pasta_temp & Replace(nomeArquivo, "VBAs", "")
        nomeArquivoFormatado = pasta_temp & Replace(nomeArquivo, "/", "")

        If nomeArquivo <> "" Then
            If Not BaixarArquivo(urlArquivo, nomeArquivoFormatado) Then
                Debug.Print "Erro ao baixar " & nomeArquivo, vbCritical
            Else
                Debug.Print iArquivo &"/"& nTotalArquivos & " - " &  nomeArquivo
            End If
        End If
    Next i
    Debug.Print "Saindo BaixarModulosViaManifestGenerico"

End Sub