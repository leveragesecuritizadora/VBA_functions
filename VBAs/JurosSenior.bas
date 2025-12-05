Attribute VB_Name = "JurosSenior"
Public Function PreencherJurosSenior( _
    Optional mes_offset As Integer = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    Dim wsAtual As Worksheet
    Dim celAtual As Range
    Dim stringBusca As String
    Dim dataBase As Date
    Dim emissao As String
    
    ' att aut das celulas a cada mudanca
    Application.Volatile True
    
    ' --- [2] Define contexto atual ---
    Set celAtual = Application.Caller
    Set wsAtual = celAtual.Parent

    ' --- [3] verifica se a coluna especificada contem uma data ---
    If IsDate(wsAtual.Cells(celAtual.Row, coluna_data).Value) Then
        dataBase = CDate(wsAtual.Cells(celAtual.Row, coluna_data).Value)
    Else
        PreencherJurosSenior = "Erro: coluna_data" & celAtual.Row & " nao contem uma data valida"
        Exit Function
    End If

    ' --- [4] Verifica se o deslocamento de mes esta dentro do intervalo ----
    If mes_offset < -12 Or mes_offset > 12 Then
        PreencherJurosSenior = "Erro: mes_offset fora do intervalo (-12 a 12)"
        Exit Function
    End If
    
    ' --- [5] Monta a string de busca ---
    emissao = Split(Application.Caller.Parent.Parent.Name, " ")(1)
    stringBusca = Format(DateSerial(Year(dataBase), Month(dataBase) + mes_offset, 1), "dd/mm/yyyy") & " - " & emissao & " - senior"
    Debug.Print "Preencher jS - busca: "; stringBusca
    Debug.Print "Preencher jS - resultado: "; BuscarLinha("Juros", 3, stringBusca)
    PreencherJurosSenior = BuscarLinha("Juros", 3, stringBusca)
    Exit Function

End Function






