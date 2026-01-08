Attribute VB_Name = "JurosSerie"
Public Function PreencherJurosSerie( _
    n_serie As Integer, _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    ' Debug.Print "PreencherJurosSerie"
    PreencherJurosSerie = ValorPrimeiroMatch(mes_offset, coluna_data, "Juros", 3, Array(n_serie, "*"))

End Function