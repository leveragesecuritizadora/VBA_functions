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

    For Each cel In planilhaDados.Range("A:A")
        If cel.Value Like id & "*" Then
            somador += planilhaDados.Cells(cel.Row, coluna_buscar).Value
            Exit For
        End If
    Next cel

    SomaValores = somador

End Function

