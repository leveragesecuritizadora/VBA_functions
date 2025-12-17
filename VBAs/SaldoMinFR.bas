Attribute VB_Name = "SaldoMinFR"
Function PreencherSaldoMinFR( _
    Optional mes_offset As Integer = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    PreencherSaldoMinFR = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "InfosFR", 3)

End Function