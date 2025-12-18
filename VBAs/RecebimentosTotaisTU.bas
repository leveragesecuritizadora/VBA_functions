Attribute VB_Name = "RecebimentosTotaisTU"
Function PreencherRecebimentosTotaisTU( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' PrintIniFuncao("R. Totais TU")
    PreencherRecebimentosTotaisTU = ImplementacaoBuscarInfosUnidades(mes_offset, coluna_data, "Recebimentos", 5)

End Function