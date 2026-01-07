Attribute VB_Name = "FBuscarInformacoesAcumuladas"
Function BuscarInformacoesAcumuladas( _
    mes_offset As Variant, _
    coluna_data As Integer, _
    planilha_dados As String, _
    coluna_dados As Integer, _
    Optional sufixo_busca As String = "" _
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

    ' Debug.Print "R" & celAtual.Row
    ' Debug.Print "C" & celAtual.Column

    dataBase = VerificaDataEOffset(wsAtual.Cells(celAtual.Row, coluna_data).Value, mes_offset)

    ' Debug.Print "Data Cascata" & dataBase

    Debug.Print "BuscarInformacoesAcumuladas: dataBase: "& dataBase

    If dataBase = False Then
        BuscarInformacoesAcumuladas = "Erro data"
        Exit Function
    End If
    
    string_busca = FormatarDataString(dataBase, mes_offset) & " - * - senior"

    Debug.print string_busca

    ' resultado = ValorAcumulado(planilha_dados, coluna_dados, string_busca)
    somador = 0
    Debug.print "chegie aqui"

    ' Set planilhaDados = ThisWorkbook.Sheets(planilha_dados)
    Debug.Print planilha_dados
    Set wb_aux = Application.Caller.Parent.Parent
    Set planilhaDados = wb.Worksheets(planilha_dados)
    If planilhaDados Is Nothing Then
        BuscarInformacoesAcumuladas = "Aba não encontrada"
        Exit Function
    Else 
        Debug.print "tudo certo com a aba"
    End If

    Debug.print "chegie aqui2"


    For Each cel In planilhaDados.Range("A1:A" &  planilhaDados.Cells(planilhaDados.Rows.Count, "A").End(xlUp).Row)
        Debug.print "Buscado: " string_busca 
        Debug.print "Val atual: " cel.Value
        Debug.print "----"
        If CStr(cel.Value) LIKE string_busca Then
            ' Debug.Print "val antes: "; somador
            ' Debug.Print "Soma Valores bateu Cel: " & Cstr(coluna_buscar) & Cstr(cel.Row); cel.Value
            ' Debug.Print "testando"
            If IsNumeric(planilhaDados.Cells(cel.Row, coluna_dados).Value) Then
                somador = somador + planilhaDados.Cells(cel.Row, coluna_dados).Value
            End If

            ' Debug.Print "val depois: "; somador
        End If
    Next cel

    resultado = somador

'    MeuPrint "Implementacao BD ", planilha_dados, " - busca: ", string_busca
'    MeuPrint "Implementacao BD ", planilha_dados, " - resultado: ", BuscarLinha(planilha_dados, coluna_dados, string_busca)


    If resultado = False Then
       ' Debug.Print "Erro"
        BuscarInformacoesAcumuladas = 0
        Exit Function
    End If

    BuscarInformacoesAcumuladas = resultado


End Function