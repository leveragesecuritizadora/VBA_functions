Attribute VB_Name = "SaldoFD"
Function PreencheSaldoFD( _
    Optional dado_historico As Variant, _
    Optional mes_desejado As Variant = False, _
    Optional mes_offset As Integer = -1, _
    Optional place_holder As Variant = "-", _
    Optional nome_fonte As String = "SaldoFD" _
) As Variant

    Dim wsBDconnected As Worksheet
    Dim wsAtual As Worksheet
    Dim celAtual As Range
    Dim dataBusca As String
    Dim linhaEncontrada As Variant
    Dim dataBase As Date
    Dim timestamp As String
    Dim colunaInfo As Integer

    ' Gera timestamp para log
    timestamp = Format(Now, "dd/mm/yyyy HH:nn:ss")

    
    On Error GoTo ErroHandler

    ' Atualiza a cada mudanaa
    Application.Volatile True

    ' --- [1] Verifica se a planilha fonte existe ---
    On Error Resume Next

    Set wsBDconnected = ThisWorkbook.Sheets(nome_fonte)
    Err.Clear
    
    On Error GoTo ErroHandler
    If wsBDconnected Is Nothing Then
        PreencheSaldoFD = "Erro: Tabela '" & nome_fonte & "' n�o existe"
        ' Debug.Print "[" & timestamp & "] Erro: Tabela '" & nome_fonte & "' n�o existe"
        Exit Function
    End If
    
    ' --- [2] Define contexto atual ---
    Set celAtual = Application.Caller
    Set wsAtual = celAtual.Parent

    ' --- [3] Verifica se a c_lula da coluna B cont_m uma data ---
    If IsDate(wsAtual.Cells(celAtual.Row, 2).Value) Then
        dataBase = CDate(wsAtual.Cells(celAtual.Row, 2).Value)
    Else
        PreencheSaldoFD = "Erro: c_lula B" & celAtual.Row & " n�o cont_m uma data v�lida"
        Exit Function
    End If

    ' --- [4] Verifica se o deslocamento de ms est� dentro do intervalo ----
    If mes_offset < -12 Or mes_offset > 12 Then
        PreencheSaldoFD = "Erro: mes_offset fora do intervalo (-12 a 12)"
        ' Debug.Print "[" & timestamp & "] Erro: mes_offset fora do intervalo (-12 a 12)"
        Exit Function
    End If
    ' --- [5] Monta a string de busca ---
    dataBusca = Format(DateSerial(Year(dataBase), Month(dataBase) + mes_offset, 1), "dd/mm/yyyy")
    ' Debug.Print "chegue 8.1"
    ' Debug.Print dataBusca
    
    ' --- [6] Busca o valor na planilha fonte ---
    linhaEncontrada = Application.Match(dataBusca, wsBDconnected.Range("B:B"), 0)
    ' Debug.Print "chegue 9: "
    ' Debug.Print linhaEncontrada
    
    If Not IsMissing(dado_historico) Then
        If Not IsEmpty(dado_historico) And dado_historico <> "" Then
            PreencheSaldoFD = dado_historico
            Exit Function
        End If
    End If
    
    ' --- [7] Define a coluna de retorno (corrigir se necess�rio) ---
    colunaInfo = 3  ' <--- ajuste aqui se for outra coluna
    
    ' --- [8] Retorna o valor encontrado ---
    PreencheSaldoFD = wsBDconnected.Cells(linhaEncontrada, colunaInfo).Value
    ' Debug.Print "SalMinFR"
    ' Debug.Print "[" & timestamp & "] Busca: " & dataBusca
    ' Debug.Print "[" & timestamp & "] Linha encontrada: " & linhaEncontrada
    ' Debug.Print "[" & timestamp & "] Valor retornado (" & wsBDconnected.Cells(linhaEncontrada, colunaInfo).Address & "): " & wsBDconnected.Cells(linhaEncontrada, colunaInfo).Value

    Exit Function

' --- [9] Tratamento gen_rico de erro inesperado ---
ErroHandler:
    PreencheSaldoFD = place_holder
    ' Call LogErroUDF("ERRO em PreencheSaldoFD: " & Err.Description)
End Function


Sub LogErroUDF(msg As String)
    ' Debug.Print msg
End Sub

