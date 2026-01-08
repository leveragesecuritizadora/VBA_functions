Attribute VB_Name = "AmortizacaoSeniorOrdinaria"
Public Function PreencherAmortizacaoSeniorOrdinaria( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' Debug.Print("amort senior odr")
    PreencherAmortizacaoSeniorOrdinaria = SomarValoresMultiplasLinhas(mes_offset, coluna_data, "Juros", 5, Array("*","senior"))

End Function