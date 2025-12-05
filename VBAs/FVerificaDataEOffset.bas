Attribute VB_Name = "FVerificaDataEOffset"
Function VerificaDataEOffset( _
    data As Variant, _
    offset As Variant _
) As Variant

    Debug.Print Now() & " - FVerificaDataEOffset: Data " & data
    Debug.Print Now() & " - FVerificaDataEOffset: offset " & offset

    ' --- [3] verifica se a coluna especificada contem uma data ---
    If Not IsDate(data) Then
        Debug.Print "Coluna nao contem data"
        VerificaDataEOffset = False
        Exit Function
    End If

    ' --- [4] Verifica se o deslocamento de mes esta dentro do intervalo ----
    If offset < -12 Or offset > 12 Then
        Debug.Print "Erro: offset fora do intervalo (-12 a 12)"
        VerificaDataEOffset = False
        Exit Function
    End If

    VerificaDataEOffset = CDate(data)
    Debug.Print Now() & " - FVerificaDataEOffset: Resultado " & VerificaDataEOffset
End Function
