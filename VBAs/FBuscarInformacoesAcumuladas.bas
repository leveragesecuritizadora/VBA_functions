Attribute VB_Name = "FBuscarInformacoesAcumuladas"
Function BuscarInformacoesAcumuladas( _
    planilha_dados As String, _
    mes_offset As Variant, _
    coluna_data As Integer, _
    coluna_dados As Integer, _
    Optional sufixo_busca As String = "" _ 
) As Variant

    Dim wsAtual As Worksheet
    Dim celAtual As Range
    Dim dataBase As Variant
    Dim resultado As Variant
    dim string_busca as Variant
    ' att aut das celulas a cada mudanca
    Application.Volatile True
    
    ' --- [2] Define contexto atual ---
    Set celAtual = Application.Caller
    Set wsAtual = celAtual.Parent

    ' Debug.Print "R" & celAtual.Row
    ' Debug.Print "C" & celAtual.Column

    dataBase = VerificaDataEOffset(wsAtual.Cells(celAtual.Row, coluna_data).Value, mes_offset)

    ' Debug.Print "Data Cascata" & dataBase

    ' Debug.Print Now() & "C: "& celAtual.Column & celAtual.Row & " - BuscarInformacoesAcumuladas: dataBase: "& dataBase

    If dataBase = False Then
        BuscarInformacoesAcumuladas = "Erro data"
        Exit Function
    End If
    
    string_busca = FormatarDataString(dataBase, mes_offset) & " - * - senior"

    ' resultado = ValorAcumulado(planilha_dados, coluna_dados, string_busca)
    dim somador as Variant
    somador = 0

    Set planilhaDados = Application.Caller.Parent.Parent.Worksheets(planilha)

    For Each cel In planilhaDados.Range("A1:A" &  planilhaDados.Cells(planilhaDados.Rows.Count, "A").End(xlUp).Row)
        Debug.print "Buscado: " string_busca 
        Debug.print "Val atual: " cel.Value
        ' Debug.print "----"
        If cel.Value LIKE string_busca Then
            ' Debug.Print "val antes: "; somador
            ' Debug.Print "Soma Valores bateu Cel: " & Cstr(coluna_buscar) & Cstr(cel.Row); cel.Value
            ' Debug.Print "testando"
            somador = somador + planilhaDados.Cells(cel.Row, coluna_buscar).Value
            ' Debug.Print "val depois: "; somador
        End If
    Next cel

    ValorAcumulado = somador

'    MeuPrint "Implementacao BD ", planilha_dados, " - busca: ", string_busca
'    MeuPrint "Implementacao BD ", planilha_dados, " - resultado: ", BuscarLinha(planilha_dados, coluna_dados, string_busca)


    If resultado = False Then
       ' Debug.Print "Erro"
        BuscarInformacoesAcumuladas = 0
        Exit Function
    End If

    BuscarInformacoesAcumuladas = resultado


End Function