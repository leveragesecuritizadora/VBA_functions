Attribute VB_Name = "Juros"
Function PreencheJuros( _
    Optional tipo_serie As String = "senior", _
    Optional dado_historico As Variant, _
    Optional mes_desejado As Variant = False, _
    Optional mes_offset As Integer = -1, _
    Optional place_holder As Variant = "-", _
    Optional coluna_juros As Variant = 7, _
    Optional nome_fonte As String = "Juros" _
) As Variant

    Dim wsJuros As Worksheet
    Dim wsAtual As Worksheet
    Dim celAtual As Range
    Dim dataBusca As String
    Dim linhaEncontrada As Variant
    Dim dataBase As Date
    
    On Error GoTo ErroHandler

    ' att aut das celulas a cada mudanca
    Application.Volatile True

    ' definindo a coluna dos juros
    If Left(tipo_serie, 6) = "senior" Then
        tipo_serie = "senior"
    ElseIf Left(tipo_serie, 11) = "subordinada" Then
        tipo_serie = "subordinada"
    Else
        PreencheJuros = "Erro: S_rie '" & tipo_serie & "' n�o existe"
        Exit Function
    End If
    
    ' --- [1] Verifica se a planilha fonte existe ---
    On Error Resume Next
    Set wsJuros = ThisWorkbook.Sheets(nome_fonte)
    On Error GoTo ErroHandler
    If wsJuros Is Nothing Then
        PreencheJuros = "Erro: Tabela '" & nome_fonte & "' n�o existe"
        Exit Function
    End If
    
    ' --- [2] Define contexto atual ---
    Set celAtual = Application.Caller
    Set wsAtual = celAtual.Parent

    ' --- [3] Verifica se a c_lula da coluna B cont_m uma data ---
    If IsDate(wsAtual.Cells(celAtual.Row, 2).Value) Then
        ' Debug.Print "mes coluna B"
        dataBase = CDate(wsAtual.Cells(celAtual.Row, 2).Value)
    Else
        PreencheJuros = "Erro: c_lula B" & celAtual.Row & " n�o cont_m uma data v�lida"
        Exit Function
    End If

    ' --- [4] Verifica se o deslocamento de ms est� dentro do intervalo ----
    If mes_offset < -12 Or mes_offset > 12 Then
        PreencheJuros = "Erro: mes_offset fora do intervalo (-12 a 12)"
        Exit Function
    End If
    
    ' --- [5] Monta a string de busca ---
    dataBusca = Format(DateSerial(Year(dataBase), Month(dataBase) + mes_offset, 1), "dd/mm/yyyy") & " - " & tipo_serie
    
    ' --- [7] Verifica se a coluna pedida _ v�lida ---
    If coluna_juros < 1 Or coluna_juros > wsJuros.Columns.Count Then
        PreencheJuros = "Erro: coluna_juros inv�lida (" & coluna_juros & ")"
        Exit Function
    End If
    
    ' --- [6] Busca o valor na planilha fonte ---
    linhaEncontrada = Application.Match(dataBusca, wsJuros.Range("D:D"), 0)
    
    If Not IsMissing(dado_historico) Then
        If Not IsEmpty(dado_historico) And dado_historico <> "" Then
            PreencheJuros = dado_historico
            Exit Function
        End If
    End If
    
    ' --- [8] Retorna o valor encontrado ---
    PreencheJuros = wsJuros.Cells(linhaEncontrada, coluna_juros).Value
    ' Debug.Print "Busca: " & dataBusca
    ' Debug.Print "Linha encontrada: " & linhaEncontrada
    ' Debug.Print "Valor na c_lula: " & wsJuros.Cells(linhaEncontrada, coluna_juros).Address & " = " & wsJuros.Cells(linhaEncontrada, coluna_juros).Value
    Exit Function

' --- [9] Tratamento gen_rico de erro inesperado ---
ErroHandler:
    PreencheJuros = "--"
End Function



