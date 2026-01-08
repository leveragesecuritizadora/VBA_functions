Attribute VB_Name = "SaldoMinFR"
Function PreencherSaldoMinFR( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' PrintIniFuncao("Sal min FR")
    PreencherSaldoMinFR = ValorPrimeiroMatch(mes_offset, coluna_data, "InfosFundos", 6)

End Function