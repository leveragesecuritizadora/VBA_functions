Attribute VB_Name = "SOrquestraAutomacaoPlanilha"

Sub OrquestraAutomacaoPlanilha()
    CriarBotaoComMacro "Atualizar Dados", "AtualizarTabelas|1", "planilha1", "Azul", 50, 50
    CriarBotaoComMacro "Gerar Planilha de Compartilhamento", "SanitizarXLSM", "planilha1", "Verde", 100, 50
    

End Sub

    