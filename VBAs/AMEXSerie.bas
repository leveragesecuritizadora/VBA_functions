Attribute VB_Name = "AMEXSerie"
Public Function PreencherAMEXSerie( _
    n_serie As Integer, _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    ' Debug.Print "PreencherAMEXSerie"
    PreencherAMEXSerie = ValorPrimeiroMatch(mes_offset, coluna_data, "Juros", 6, Array(n_serie, "*"))

End Function