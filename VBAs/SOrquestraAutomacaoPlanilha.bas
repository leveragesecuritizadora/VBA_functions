Attribute VB_Name = "SOrquestraAutomacaoPlanilha"

Sub OrquestraAutomacaoPlanilha()
    Dim id As Integer
    id = IDEmissao()

    LimparTerminal "Automação Planilha - ID: " & id

    CriarBotaoComMacro "Atualizar Dados", "AtualizarTabelas|" & id, "planilha1", "Azul", 250, 50
    CriarBotaoComMacro "Gerar Planilha de Compartilhamento", "SanitizarXLSM", "planilha1", "Verde", 500, 50

    AtualizarTabelas(id)
    

End Sub

    