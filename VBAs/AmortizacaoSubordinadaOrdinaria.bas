Attribute VB_Name = "AmortizacaoSubordinadaOrdinaria"
Public Function PreencherAmortizacaoSubordinadaOrdinaria( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' Debug.Print("amort ord sub")
    PreencherAmortizacaoSubordinadaOrdinaria = SomarValoresMultiplasLinhas(mes_offset, coluna_data, "Juros", 5, Array("*","subordinada"))

End Function