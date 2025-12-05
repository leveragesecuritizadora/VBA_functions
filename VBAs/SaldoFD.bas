Attribute VB_Name = "SaldoFD"
Function PreencherSaldoFD( _
    Optional mes_offset As Integer = -1, _
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

    ' Debug.Print Now() & "C: "& celAtual.Column & celAtual.Row & " - PreencherSaldoFD: dataBase: "& dataBase

    If dataBase = False Then
        PreencherSaldoFD = "Erro data"
        Exit Function
    End If
    
    ' --- [5] Monta a string de busca ---
    emissao = Split(Application.Caller.Parent.Parent.Name, " ")(1)
    stringBusca = Format(DateSerial(Year(dataBase), Month(dataBase) + mes_offset, 1), "dd/mm/yyyy") & " - " & emissao
    resultado = BuscarLinha("SaldoFD", 2, stringBusca)

    ' Debug.Print "Preencher saldoFR - busca: "; stringBusca
    ' Debug.Print "Preencher saldoFR - resultado: "; BuscarLinha("SaldoFD", 2, stringBusca)


    If resultado = False Then
        PreencherSaldoFD = 0
        Exit Function
    End If

    PreencherSaldoFD = resultado
    ' PreencherSaldoFD = "OK"


End Function