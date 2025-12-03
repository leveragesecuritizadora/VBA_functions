Attribute VB_Name = "Recebimentos"
Function PreencheRecebimento( _
    Optional unidade As String = "Unidade", _
    Optional tipo_recebimento As String = "total", _
    Optional dado_historico As Variant, _
    Optional mes_desejado As Variant = False, _
    Optional mes_offset As Integer = -1, _
    Optional place_holder As Variant = "-", _
    Optional nome_fonte As String = "Recebimentos" _
) As Variant

    Dim wsRecebimentos As Worksheet
    Dim wsAtual As Worksheet
    Dim celAtual As Range
    Dim dataBusca As String
    Dim linhaEncontrada As Variant
    Dim dataBase As Date
    Dim colunaRecebimento As Integer
    
    On Error GoTo ErroHandler

    ' att aut das celulas a cada mudanca
    Application.Volatile True
    
    Select Case tipo_recebimento
        Case "total"
            colunaRecebimento = 8
        Case "antecipado"
            colunaRecebimento = 6
        Case Else
            PreencheRecebimento = "Erro: '" & tipo_recebimento & "' n�o existe"
            Exit Function
    End Select


    ' --- [1] Verifica se a planilha fonte existe ---
    On Error Resume Next
    Set wsRecebimentos = ThisWorkbook.Sheets(nome_fonte)
    On Error GoTo ErroHandler
    If wsRecebimentos Is Nothing Then
        PreencheRecebimento = "Erro: Tabela '" & nome_fonte & "' n�o existe"
        Exit Function
    End If
    
    ' --- [2] Define contexto atual ---
    Set celAtual = Application.Caller
    Set wsAtual = celAtual.Parent

    ' --- [3] Verifica se a c_lula da coluna B cont_m uma data ---
    If IsDate(wsAtual.Cells(celAtual.Row, 2).Value) Then
        dataBase = CDate(wsAtual.Cells(celAtual.Row, 2).Value)
    Else
        PreencheRecebimento = "Erro: c_lula B" & celAtual.Row & " n�o cont_m uma data v�lida"
        Exit Function
    End If

    ' --- [4] Verifica se o deslocamento de ms est� dentro do intervalo ----
    If mes_offset < -12 Or mes_offset > 12 Then
        PreencheRecebimento = "Erro: mes_offset fora do intervalo (-12 a 12)"
        Exit Function
    End If
    
    ' --- [5] Monta a string de busca ---
    dataBusca = Format(DateSerial(Year(dataBase), Month(dataBase) + mes_offset, 1), "dd/mm/yyyy") & " - " & unidade
    
    ' --- [7] Verifica se a coluna pedida _ v�lida ---
    If colunaRecebimento < 1 Or colunaRecebimento > wsRecebimentos.Columns.Count Then
        PreencheRecebimento = "Erro: colunaRecebimento inv�lida (" & colunaRecebimento & ")"
        Exit Function
    End If
    
    ' --- [6] Busca o valor na planilha fonte ---
    linhaEncontrada = Application.Match(dataBusca, wsRecebimentos.Range("D:D"), 0)
    
    If Not IsMissing(dado_historico) Then
        If Not IsEmpty(dado_historico) And dado_historico <> "" Then
            PreencheRecebimento = dado_historico
            Exit Function
        End If
    End If
    
    ' --- [8] Retorna o valor encontrado ---
    PreencheRecebimento = wsRecebimentos.Cells(linhaEncontrada, colunaRecebimento).Value
    ' Debug.Print "Busca: " & dataBusca
    ' Debug.Print "Linha encontrada: " & linhaEncontrada
    ' Debug.Print "Valor na c_lula: " & wsRecebimentos.Cells(linhaEncontrada, colunaRecebimento).Address & " = " & wsRecebimentos.Cells(linhaEncontrada, colunaRecebimento).Value

    Exit Function

' --- [9] Tratamento gen_rico de erro inesperado ---
ErroHandler:
    PreencheRecebimento = "--"
End Function



