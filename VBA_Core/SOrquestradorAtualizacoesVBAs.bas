Attribute VB_Name = "SOrquestradorAtualizacoesVBAs"

Public Sub OrquestradorAtualizacoesVBAs()
    LimparTerminal "Rodando orquestrador 3.0"
    On Error GoTo Fim

    Application.ScreenUpdating = False

    Call ApagarModulos
    ' Debug.Print "rodou ApagarModulos"
    Call BaixarModulosViaManifest
    ' Debug.Print "rodou BaixarModulosViaManifest"
    Call ImportarModulos
    ' Debug.Print "rodou ImportarModulos"
    OrquestradorAutomacaoPlanilha

    Application.ScreenUpdating = True
    MsgBox Now & " Projeto VBA atualizado com sucesso!", vbInformation

    Fim:
        Application.ScreenUpdating = True
End Sub

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
    Dim vbComp As Object

    LimparTerminal "Apagando módulos antigos (modo agressivo)"

    For Each vbComp In ThisWorkbook.VBProject.VBComponents

        ' Apenas módulos padrão (.bas)
        If vbComp.Type = 1 Then

            ' Nunca apagar o bootloader nem o orquestrador (mesmo se tiver sufixo)
            If Not vbComp.Name Like "Bootloader*" _
               And Not vbComp.Name Like "SOrquestradorAtualizacoesVBAs*" Then

                Debug.Print "Removendo módulo: "; vbComp.Name
                ThisWorkbook.VBProject.VBComponents.Remove vbComp

            Else
                Debug.Print "Preservado: "; vbComp.Name
            End If

        End If

    Next vbComp

End Sub


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

Private Sub ImportarModulos()
    Dim pasta As String
    Dim arquivo As String

    ' Debug.Print "Dentro ImportarModulos"

    pasta = Environ("TEMP") & "\vba\"
    arquivo = Dir(pasta & "*.bas")
    ' Debug.Print "meio ImportarModulos"

    Do While arquivo <> ""
        If arquivo <> "OrquestradorAtualizacoesVBAs.bas" Then
            ThisWorkbook.VBProject.VBComponents.Import pasta & arquivo
        End If
        arquivo = Dir
    Loop
    ' Debug.Print "saindo ImportarModulos"

End Sub

Private Function BaixarArquivo(url As String, destino As String) As Boolean
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

Private Sub OrquestradorAutomacaoPlanilha()
    Dim id As Integer
    id = IDEmissao()

    LimparTerminal "Automação Planilha - ID: " & id

    SOrquestradorAtualizacoesVBAs.CriarBotaoComMacro "Atualizar Dados", "AtualizarTabelas|" & id, "Ordem de Pagamento Consolidado", "Azul", 250, 50
    SOrquestradorAtualizacoesVBAs.CriarBotaoComMacro "Atualizar Módulos", "RodarBootloader", "Ordem de Pagamento Consolidado", "Verde", 350, 50

    AtualizarTabelas(id)
End Sub

Private Sub LimparTerminal(mensagem As String) 
    Debug.Print String(80, "=")
    Debug.Print Now & " " & mensagem
    Debug.Print String(80, "=")
End Sub

Sub CriarBotaoComMacro( _
    texto_botao As String, _
    funcao_argumento As String, _
    nome_aba As String, _
    cor_botao As String, _
    Optional left_pos As Double = 50, _
    Optional top_pos As Double = 50 _
)

    Dim ws As Worksheet
    Dim btn As Shape
    Dim larguraMin As Double
    Dim padding As Double

    padding = 20
    larguraMin = 80

    Set ws = ThisWorkbook.Sheets(nome_aba)

    ' Se já existir, remove (evita duplicar)
    If BotaoExiste(ws, funcao_argumento) Then
        Debug.Print "botao " & funcao_argumento & " já existe, deletando botão..."
        ws.Shapes(funcao_argumento).Delete
    End If

    ' Cria o botão
    Set btn = ws.Shapes.AddShape( _
        Type:=msoShapeRoundedRectangle, _
        Left:=left_pos, _
        Top:=top_pos, _
        Width:=larguraMin, _
        Height:=35 _
    )

    With btn

        .Name = funcao_argumento
        .TextFrame2.TextRange.Text = texto_botao

        ' Fonte
        With .TextFrame2.TextRange.Font
            .Size = 11
            .Bold = msoTrue
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With

        ' Centralização
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .MarginLeft = 5
            .MarginRight = 5
            .MarginTop = 2
            .MarginBottom = 2
        End With

        ' Cor do botão
        .Fill.ForeColor.RGB = CorPorNome(cor_botao)

        ' Remove borda
        ' .Line.Visible = msoFalse

        ' Autoajuste de largura pelo texto
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText

        ' Garante largura mínima + padding
        If .Width < larguraMin Then .Width = larguraMin
        .Width = .Width + padding

        ' Vincula a macro
        .OnAction = "ChamaFuncaoCmArgumento"

    End With

    LimparTerminal "Botão '" & texto_botao & "' criado com sucesso"

End Sub

Private Function CorPorNome(nomeCor As String) As Long

    Select Case LCase(nomeCor)
        Case "azul"
            CorPorNome = RGB(0, 112, 192)
        Case "cinza"
            CorPorNome = RGB(128, 128, 128)
        Case "verde"
            CorPorNome = RGB(0, 176, 80)
        Case "vermelho"
            CorPorNome = RGB(192, 0, 0)
        Case "laranja"
            CorPorNome = RGB(237, 125, 49)
        Case "preto"
            CorPorNome = RGB(0, 0, 0)
        Case Else
            ' cor padrão
            CorPorNome = RGB(0, 112, 192)
    End Select

End Function

Private Function BotaoExiste(ws As Worksheet, nomeShape As String) As Boolean
    On Error Resume Next
    BotaoExiste = Not ws.Shapes(nomeShape) Is Nothing
    On Error GoTo 0
End Function

Public Sub ChamaFuncaoCmArgumento()
    Dim nomeBotao As String
    Dim partes() As String
    Dim parametro As String
    Dim funcao As String

    nomeBotao = Application.Caller

    Debug.Print "nomeBotao: "; nomeBotao

    If nomeBotao LIKE "*|*" Then

        partes = Split(nomeBotao, "|")

        funcao = partes(0)
        parametro = CInt(partes(1))

        Debug.Print funcao, parametro

        If funcao = "AtualizarTabelas" Then
            AtualizarTabelas(parametro)
        End If
    
    Else
        If nomeBotao = "RodarBootloader" Then
            Call RodarBootloader
        End If
    
    End If
        
End Sub