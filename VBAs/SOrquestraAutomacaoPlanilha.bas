Attribute VB_Name = "SOrquestraAutomacaoPlanilha"

Sub OrquestraAutomacaoPlanilha()
    Dim id As Integer
    id = IDEmissao()

    LimparTerminal "Automação Planilha - ID: " & id

    CriarBotaoComMacro "Atualizar Dados", "AtualizarTabelas|" & id, "planilha1", "Azul", 50, 50
    CriarBotaoComMacro "Gerar Planilha de Compartilhamento", "SanitizarXLSM", "planilha1", "Verde", 300, 50
    

End Sub

    