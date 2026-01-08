Attribute VB_Name = "AMEXSubordinada"
Public Function PreencherAmexSubordinada( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' Debug.Print("AMEXSub")
    PreencherAmexSubordinada = SomarValoresMultiplasLinhas(mes_offset, coluna_data, "Juros", 6, Array("*","subordinada"))

End Function