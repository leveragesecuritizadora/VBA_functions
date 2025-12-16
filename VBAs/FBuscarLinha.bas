Attribute VB_Name = "FBuscarLinha"
Function BuscarLinha( _
    planilha As Variant, _
    coluna_buscar As Variant, _
    id As Variant, _
    Optional log As Variant = False _
) As Variant
    Dim wb As Workbook
    Dim planilhaDados As Worksheet
    Dim linhaEncontrada As Variant

    Set planilhaDados = Application.Caller.Parent.Parent.Worksheets(planilha)

    linhaEncontrada = Application.Match(id, planilhaDados.Range("A:A"), 0)

    ' Debug.Print "FBuscarLInha - ID: "; id
    ' Debug.Print "FBuscarLinha - LE: "; linhaEncontrada


    If IsError(linhaEncontrada) Then
        If log Then
            Dim diaHora As Date
            diaHora = Now()
            ' Debug.Print diaHora & " - BuscarLinha: Celula com identificador (" & id & ") não encontrada em " & planilha 
        End If
        BuscarLinha = False
        Exit Function
    End If

    BuscarLinha = planilhaDados.Cells(linhaEncontrada, coluna_buscar).Value

    ' Debug.Print "FBuscarLinha - valor encontrada: "; BuscarLinha
End Function
