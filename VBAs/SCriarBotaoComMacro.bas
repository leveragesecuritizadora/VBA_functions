Attribute VB_Name = "SCriarBotaoComMacro"
Sub CriarBotaoComMacro( _
    texto_botao As String, _
    funcao_botao As String, _
    nome_aba As String _
)

    Dim ws As Worksheet
    Dim btn As Shape

    Set ws = ThisWorkbook.Sheets(nome_aba) ' <-- altere para sua aba

    ' Cria o botão (shape)
    Set btn = ws.Shapes.AddShape( _
        Type:=msoShapeRoundedRectangle, _
        Left:=50, _
        Top:=50, _
        Width:=150, _
        Height:=40 _
    )

    ' Configura o visual
    With btn
        .Name = "btnAtualizar"
        .TextFrame2.TextRange.Text = texto_botao
        .Fill.ForeColor.RGB = RGB(0, 112, 192)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Bold = msoTrue

        ' Vincula a macro
        .OnAction = funcao_botao
    End With

    LimparTerminal "Botão " & texto_botao & " criado" 

End Sub
