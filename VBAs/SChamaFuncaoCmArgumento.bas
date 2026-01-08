Attribute VB_Name = "SChamaFuncaoCmArgumento"
Sub ChamaFuncaoCmArgumento()
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

        If nomeBotao = "AtualizarModulos" Then
            Debug.Print "AtualizarModulos()"
        End If
    
    End If
        
End Sub
