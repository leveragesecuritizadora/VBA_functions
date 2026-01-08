Attribute VB_Name = "FValorPrimeiroMatch"
Function ValorPrimeiroMatch( _
    mes_offset As Variant, _
    coluna_data As Integer, _
    planilha_dados As String, _
    coluna_dados As Integer, _
    Optional sufixo_busca As Variant _
) As Variant

    Dim wsAtual As Worksheet
    Dim celAtual As Range
    Dim dataBase As Variant
    Dim string_busca as Variant
    Dim planilhaDados As Worksheet
    Dim cel As Range
    Dim linhaEncontrada as Variant

    ' att aut das celulas a cada mudanca
    Application.Volatile True
    
    ' --- [2] Define contexto atual ---
    Set celAtual = Application.Caller
    Set wsAtual = celAtual.Parent

    dataBase = VerificaDataEOffset(wsAtual.Cells(celAtual.Row, coluna_data).Value, mes_offset)

    If dataBase = False Then
        ValorPrimeiroMatch = "Erro data"
        Exit Function
    End If
    
    If IsArray(sufixo_busca) Then
        string_busca = FormatarDataString(dataBase, mes_offset) &  " - " & Join(sufixo_busca, " - ")
    Else 
        string_busca = FormatarDataString(dataBase, mes_offset)
    End If

    Debug.print string_busca

    somador = 0
    Set planilhaDados = ThisWorkbook.Worksheets(planilha_dados)
    If planilhaDados Is Nothing Then
        ValorPrimeiroMatch = "Aba não encontrada"
        Exit Function
    End If

    linhaEncontrada = Application.Match(string_busca, planilhaDados.Range("A:A"), 0)

    If IsError(linhaEncontrada) Then
        Debug.print "ValorPrimeiroMatch - Linha não encontrada"
        ValorPrimeiroMatch = 0
        Exit Function
    Else
        Debug.print "Linha encontrada"
        Debug.Print "ValorPrimeiroMatch - ID: "; string_busca
        Debug.Print "ValorPrimeiroMatch - LE: "; linhaEncontrada
    End If

    Debug.print planilhaDados.Cells(linhaEncontrada, coluna_dados).Value

    ValorPrimeiroMatch = planilhaDados.Cells(linhaEncontrada, coluna_dados).Value

End Function