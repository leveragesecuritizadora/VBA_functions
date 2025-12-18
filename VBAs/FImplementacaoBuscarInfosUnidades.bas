Attribute VB_Name = "FImplementacaoBuscarInfosUnidades"
Function ImplementacaoBuscarInfosUnidades( _
    mes_offset As Variant, _
    coluna_data As Integer, _
    planilha_dados As String, _
    coluna_dados As Integer _
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

    ' Debug.Print Now() & "C: "& celAtual.Column & celAtual.Row & " - ImplementacaoBuscarInfosUnidades: dataBase: "& dataBase

    If dataBase = False Then
        ImplementacaoBuscarInfosUnidades = "Erro data"
        Exit Function
    End If
    
    ' --- [5] Monta a string de busca ---

    emissao = NomeEmissao()
    stringBusca = Format(DateSerial(Year(dataBase), Month(dataBase) + mes_offset, 1), "dd/mm/yyyy") & " - " & emissao
    resultado = SomaValores(planilha_dados, coluna_dados, stringBusca)

    ' Debug.Print "Val. tds Unidades - busca: "; stringBusca
    ' Debug.Print "Val. tds Unidades - resultado: "; SomaValores(planilha_dados, coluna_dados, stringBusca)


    If resultado = False Then
        ImplementacaoBuscarInfosUnidades = 0
        Exit Function
    End If

    ImplementacaoBuscarInfosUnidades = resultado

End Function