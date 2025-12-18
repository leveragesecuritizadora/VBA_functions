Attribute VB_Name = "SaldoMinFD"
Function PreencherSaldoMinFD( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    PreencherSaldoMinFD = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "InfosFundos", 3)

End Function