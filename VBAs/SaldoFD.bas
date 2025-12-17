Attribute VB_Name = "SaldoFD"
Function PreencherSaldoFD( _
    Optional mes_offset As Integer = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    PreencherSaldoFD = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "InfosFD", 2)


End Function