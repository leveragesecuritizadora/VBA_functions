Attribute VB_Name = "FImplementacaoBuscarInfosEmissao"
Function ImplementacaoBuscarInfosEmissao( _
    mes_offset As Integer, _
    coluna_data As Integer, _
    planilha_dados As String, _
    coluna_dados As Integer, _
    Optional nome_unidade As String = "" _ 
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

    Debug.Print "R" & dataBase

    ' Debug.Print Now() & "C: "& celAtual.Column & celAtual.Row & " - ImplementacaoBuscarInfosEmissao: dataBase: "& dataBase

    If dataBase = False Then
        ImplementacaoBuscarInfosEmissao = "Erro data"
        Exit Function
    End If
    
    ' ' --- [5] Monta a string de busca ---
    emissao = NomeEmissao()
    If Len(Trim(nome_unidade)) > 0 Then
        stringBusca = FormatarDataString(dataBase, mes_offset) & " - " & emissao & " - " & nome_unidade ' Info de uniade
    Else
        stringBusca = FormatarDataString(dataBase, mes_offset) & " - " & emissao ' Info de emissao
    End If

    resultado = BuscarLinha(planilha_dados, coluna_dados, stringBusca)

    Debug.Print "Implementacao BD - busca: "; stringBusca
    Debug.Print "Implementacao BD - resultado: "; BuscarLinha(planilha_dados, coluna_dados, stringBusca)


    If resultado = False Then
        ImplementacaoBuscarInfosEmissao = 0
        Exit Function
    End If

    ImplementacaoBuscarInfosEmissao = resultado


End Function