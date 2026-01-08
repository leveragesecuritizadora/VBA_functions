Attribute VB_Name = "FPreencherJurosSerie"
Public Function PreencherJurosSerie( _
    n_serie As Integer, _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    Debug.Print "JS"
    PreencherJurosSerie = SomarValoresMultiplasLinhas(mes_offset, coluna_data, "Juros", 3, Array("*", "*"))

End Function