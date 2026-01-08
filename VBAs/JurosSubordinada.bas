Attribute VB_Name = "JurosSubordinada"
Public Function PreencherJurosSubordinada( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    ' Debug.Print "JSub"
    PreencherJurosSubordinada = SomarValoresMultiplasLinhas(mes_offset, coluna_data, "Juros", 3, Array("*", "subordinada"))

End Function






