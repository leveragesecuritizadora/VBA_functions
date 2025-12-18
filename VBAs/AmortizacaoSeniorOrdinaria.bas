Attribute VB_Name = "AmortizacaoSeniorOrdinaria"
Public Function PreencherAmortizacaoSeniorOrdinaria( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' PrintIniFuncao("amort senior odr")
    PreencherAmortizacaoSeniorOrdinaria = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "Juros", 5, "senior")

End Function