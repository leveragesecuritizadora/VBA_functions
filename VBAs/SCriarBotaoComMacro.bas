Attribute VB_Name = "SCriarBotaoComMacro"

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