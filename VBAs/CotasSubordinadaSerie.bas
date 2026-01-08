Attribute VB_Name = "CotasSubordinada"
Function PreencherCotasSubordinada( _
    n_serie As Integer, _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' Debug.Print "Cotas Subordinada"
    PreencherCotasSubordinada = ValorPrimeiroMatch(mes_offset, coluna_data, "Juros", 2, Array(n_serie, "subordinada"))

End Function