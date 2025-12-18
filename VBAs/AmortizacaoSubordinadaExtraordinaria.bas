Attribute VB_Name = "AMEXSubordinada"
Public Function PreencherAmexSubordinada( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    Dim wsAtual As Worksheet
    Dim celAtual As Range
    Dim stringBusca As String
    Dim dataBase As Variant
    Dim emissao As String
    Dim resultado As Variant
    
    ' att aut das celulas a cada mudanca
    Application.Volatile True
    
    ' --- [2] Define contexto atual ---
    Set celAtual = Application.Caller
    Set wsAtual = celAtual.Parent

    ' Debug.Print "R" & celAtual.Row
    ' Debug.Print "C" & celAtual.Column

    dataBase = VerificaDataEOffset(wsAtual.Cells(celAtual.Row, coluna_data).Value, mes_offset)

    ' Debug.Print "R" & dataBase

    ' Debug.Print Now() & "C: "& celAtual.Column & celAtual.Row & " - PreencherAmexSubordinada: dataBase: "& dataBase

    If dataBase = False Then
        PreencherAmexSubordinada = "Erro data"
        Exit Function
    End If
    
    ' --- [5] Monta a string de busca ---
    Dim nomePlanilha As String
    nomePlanilha = Application.Caller.Parent.Parent.Name

    emissao = nomePlanilha
    emissao = Replace(emissao, "CRI ", "")
    emissao = Replace(emissao, " - Cascata.Automatizada.VBA.xlsm", "")
    stringBusca = Format(DateSerial(Year(dataBase), Month(dataBase) + mes_offset, 1), "dd/mm/yyyy") & " - " & emissao & " - subordinada"
    resultado = BuscarLinha("Juros", 6, stringBusca)

    ' Debug.Print "Preencher jS - busca: "; stringBusca
    ' Debug.Print "Preencher jS - resultado: "; BuscarLinha("Juros", 3, stringBusca)


    If resultado = False Then
        PreencherAmexSubordinada = 0
        Exit Function
    End If

    PreencherAmexSubordinada = resultado

End Function