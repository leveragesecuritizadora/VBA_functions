Attribute VB_Name = "PMTSubordinada"
Public Function PreencherPMTSubordinada( _
    Optional mes_offset As Integer = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    Dim wsAtual As Worksheet
    Dim celAtual As Range
    Dim stringBusca As String
    Dim dataBase As Variant
    Dim emissao As String
    Dim resultado As Variant

    ' Verificacao para PMTs futuras
    If Not (mes_offset = -1) Then
        mes_offset = mes_offset - 1
        ' Debug.Print "PreencherPMTSubordinada - offset transformado: "; mes_offset
    End If
    
    ' att aut das celulas a cada mudanca
    Application.Volatile True
    
    ' --- [2] Define contexto atual ---
    Set celAtual = Application.Caller
    Set wsAtual = celAtual.Parent

    ' Debug.Print "R" & celAtual.Row
    ' Debug.Print "C" & celAtual.Column

    dataBase = VerificaDataEOffset(wsAtual.Cells(celAtual.Row, coluna_data).Value, mes_offset)

    ' Debug.Print "R" & dataBase

    ' Debug.Print Now() & "C: "& celAtual.Column & celAtual.Row & " - PreencherPMTSubordinada: dataBase: "& dataBase

    If dataBase = False Then
        PreencherPMTSubordinada = "Erro data"
        Exit Function
    End If
    
    ' --- [5] Monta a string de busca ---
    emissao = Split(Application.Caller.Parent.Parent.Name, " ")(1)
    stringBusca = Format(DateSerial(Year(dataBase), Month(dataBase) + mes_offset, 1), "dd/mm/yyyy") & " - " & emissao & " - subordinada"
    resultado = BuscarLinha("Juros", 7, stringBusca)

    ' Debug.Print "Preencher jS - busca: "; stringBusca
    ' Debug.Print "Preencher jS - resultado: "; BuscarLinha("Juros", 3, stringBusca)


    If resultado = False Then
        PreencherPMTSubordinada = 0
        Exit Function
    End If

    PreencherPMTSubordinada = resultado

End Function






