Attribute VB_Name = "JurosSenior"
Public Function PreencherJurosSenior( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    Debug.Print "JS"
    PreencherJurosSenior = SomarValoresMultiplasLinhas(mes_offset, coluna_data, "Juros", 3, Array("*", "senior"))

End Function






