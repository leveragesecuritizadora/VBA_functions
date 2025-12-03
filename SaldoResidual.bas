Attribute VB_Name = "SaldoResidual"
Function PreencheSaldoResidual( _
    Optional tipo_serie As String = "senior", _
    Optional dado_historico As Variant, _
    Optional mes_desejado As Variant = False, _
    Optional mes_offset As Integer = -1, _
    Optional place_holder As Variant = "-", _
    Optional nome_fonte As String = "SaldoResidual" _
) As Variant

    Dim wsVP As Worksheet
    Dim wsAtual As Worksheet
    Dim celAtual As Range
    Dim dataBusca As String
    Dim linhaEncontrada As Variant
    Dim dataBase As Date
    Dim colunaInfo As Integer

    colunaInfo = 3
    
    On Error GoTo ErroHandler

    ' att aut das celulas a cada mudanca
    Application.Volatile True

    ' --- [1] Verifica se a planilha fonte existe ---
    ' Debug.Print "chegeu 1";
    On Error Resume Next
    Set wsVP = ThisWorkbook.Sheets(nome_fonte)
    On Error GoTo ErroHandler
    If wsVP Is Nothing Then
        PreencheSaldoResidual = "Erro: Tabela '" & nome_fonte & "' n�o existe"
        Exit Function
    End If
    
    ' --- [2] Define contexto atual ---
    ' Debug.Print "chegeu 2";
    Set celAtual = Application.Caller
    Set wsAtual = celAtual.Parent

    ' --- [3] Verifica se a c_lula da coluna B cont_m uma data ---
    If IsDate(wsAtual.Cells(celAtual.Row, 2).Value) Then
        dataBase = CDate(wsAtual.Cells(celAtual.Row, 2).Value)
    Else
        PreencheSaldoResidual = "Erro: c_lula B" & celAtual.Row & " n�o cont_m uma data v�lida"
        Exit Function
    End If

    ' --- [4] Verifica se o deslocamento de ms est� dentro do intervalo ----
    ' Debug.Print "chegeu 4";
    If mes_offset < -12 Or mes_offset > 12 Then
        PreencheSaldoResidual = "Erro: mes_offset fora do intervalo (-12 a 12)"
        Exit Function
    End If
    
    ' --- [5] Monta a string de busca ---
    ' Debug.Print "chegeu 5";
    dataBusca = Format(DateSerial(Year(dataBase), Month(dataBase) + mes_offset, 1), "dd/mm/yyyy") & " - " & tipo_serie
    
    ' --- [6] Busca o valor na planilha fonte ---
    ' Debug.Print "chegeu 6";
    linhaEncontrada = Application.Match(dataBusca, wsVP.Range("B:B"), 0)
    
    If Not IsMissing(dado_historico) Then
        If Not IsEmpty(dado_historico) And dado_historico <> "" Then
            PreencheSaldoResidual = dado_historico
            Exit Function
        End If
    End If
    
    ' --- [8] Retorna o valor encontrado ---
    ' Debug.Print "chegeu 8 residual"; linhaEncontrada
    PreencheSaldoResidual = wsVP.Cells(linhaEncontrada, colunaInfo).Value
    ' Debug.Print "Busca: " & dataBusca
    ' Debug.Print "Linha encontrada: " & linhaEncontrada
    ' Debug.Print "Valor na c_lula: " & wsVP.Cells(linhaEncontrada, colunaInfo).Address & " = " & wsVP.Cells(linhaEncontrada, colunaInfo).Value

    Exit Function

' --- [9] Tratamento gen_rico de erro inesperado ---
ErroHandler:
    PreencheSaldoResidual = "--"
End Function

