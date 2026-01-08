Attribute VB_Name = "AMEXSenior"
Public Function PreencherAmexSenior( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' PrintIniFuncao("AMEXSenior")
    PreencherAmexSenior = SomarValoresMultiplasLinhas(mes_offset, coluna_data, "Juros", 6, Array("*","senior"))

End Function