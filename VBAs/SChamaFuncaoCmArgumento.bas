Attribute VB_Name = "SChamaFuncaoCmArgumento"
Sub ChamaFuncaoCmArgumento()
    Dim nomeBotao As String
    Dim partes() As String
    Dim parametro As String
    Dim funcao As String

    nomeBotao = Application.Caller
    partes = Split(nomeBotao, "|")

    funcao - partes(0)
    parametro = CLng(parametro)(partes(1))

    If funcao = "AtualizarTabelas" Then
        AtualizarTabelas(parametro)
    End If



End Sub
