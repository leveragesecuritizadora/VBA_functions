Attribute VB_Name = "FSomaValores"
Function SomaValores( _
    planilha As Variant, _
    coluna_buscar As Variant, _
    id As Variant, _
    Optional log As Variant = False _
) As Variant
    Dim wb As Workbook
    Dim planilhaDados As Worksheet
    Dim linhaEncontrada As Variant
    Dim cel As Range
    Dim somador As Variant

    somador = 0

    Set planilhaDados = Application.Caller.Parent.Parent.Worksheets(planilha)

    For Each cel In planilhaDados.Range("A1:A" &  planilhaDados.Cells(planilhaDados.Rows.Count, "A").End(xlUp).Row)
        ' Debug.print cel.Value
        ' Debug.print "----"
        If Left(cel.Value, Len(id)) = id Then
            ' Debug.Print "val antes: "; somador
            ' Debug.Print "bateu Cel: " & Cstr(coluna_buscar) & Cstr(cel.Row); cel.Value
            ' Debug.Print "testando"
            somador = somador + planilhaDados.Cells(cel.Row, coluna_buscar).Value
            ' Debug.Print "val depois: "; somador
        End If
    Next cel

    SomaValores = somador

End Function

