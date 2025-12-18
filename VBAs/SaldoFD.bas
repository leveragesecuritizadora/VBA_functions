Attribute VB_Name = "SaldoFD"
Function PreencherSaldoFD( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    PrintIniFuncao("SaldoFD")
    PreencherSaldoFD = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "InfosFundos", 2)


End Function