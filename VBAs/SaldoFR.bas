Attribute VB_Name = "SaldoFR"
Function PreencherSaldoFR( _
    Optional mes_offset As Integer = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    PreencherSaldoFR = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "InfosFR", 2)



End Function