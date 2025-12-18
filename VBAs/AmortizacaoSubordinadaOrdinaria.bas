Attribute VB_Name = "AmortizacaoSubordinadaOrdinaria"
Public Function PreencherAmortizacaoSubordinadaOrdinaria( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' PrintIniFuncao("amort ord sub")
    PreencherAmortizacaoSubordinadaOrdinaria = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "Juros", 5, "subordinada")

End Function