Attribute VB_Name = "AmortizacaoOrdinariaSerie"
Public Function PreencherAmortizacaoOrdinariaSerie( _
    n_serie As Integer, _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    ' Debug.Print "PreencherAmortizacaoOrdinariaSerie"
    PreencherAmortizacaoOrdinariaSerie = ValorPrimeiroMatch(mes_offset, coluna_data, "Juros", 5, Array(n_serie, "*"))

End Function