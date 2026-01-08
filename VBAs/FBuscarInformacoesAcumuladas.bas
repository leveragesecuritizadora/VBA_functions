Attribute VB_Name = "FSomarValoresMultiplasLinhas"
Function SomarValoresMultiplasLinhas( _
    mes_offset As Variant, _
    coluna_data As Integer, _
    planilha_dados As String, _
    coluna_dados As Integer, _
    Optional sufixo_busca As Variant _
) As Variant

    Dim wsAtual As Worksheet
    Dim celAtual As Range
    Dim dataBase As Variant
    Dim resultado As Variant
    Dim string_busca as Variant
    Dim planilhaDados As Worksheet
    Dim cel As Range
    Dim somador as Variant
    Dim wb_aux As Workbook

    ' att aut das celulas a cada mudanca
    Application.Volatile True
    
    ' --- [2] Define contexto atual ---
    Set celAtual = Application.Caller
    Set wsAtual = celAtual.Parent

    dataBase = VerificaDataEOffset(wsAtual.Cells(celAtual.Row, coluna_data).Value, mes_offset)

    If dataBase = False Then
        SomarValoresMultiplasLinhas = "Erro data"
        Exit Function
    End If
    
    If UBound(sufixo_busca) > 0 Then
        string_busca = FormatarDataString(dataBase, mes_offset) &  " - " & Join(sufixo_busca, " - ")
    Else 
        string_busca = FormatarDataString(dataBase, mes_offset)
    End If

    Debug.print string_busca

    somador = 0
    Set planilhaDados = ThisWorkbook.Worksheets(planilha_dados)
    If planilhaDados Is Nothing Then
        SomarValoresMultiplasLinhas = "Aba não encontrada"
        Exit Function
    End If

    For Each cel In planilhaDados.Range("A1:A" &  planilhaDados.Cells(planilhaDados.Rows.Count, "A").End(xlUp).Row)
        Debug.print "----"
        Debug.print "Buscado: " string_busca 
        Debug.print "Val atual: " cel.Value
        If CStr(cel.Value) LIKE string_busca Then
            If IsNumeric(planilhaDados.Cells(cel.Row, coluna_dados).Value) Then
                Debug.Print "Bateu, somando"
                somador = somador + planilhaDados.Cells(cel.Row, coluna_dados).Value
            End If
        End If
    Next cel

    resultado = somador

    If resultado = False Then
       ' Debug.Print "Erro"
        SomarValoresMultiplasLinhas = 0
        Exit Function
    End If

    SomarValoresMultiplasLinhas = resultado

End Function