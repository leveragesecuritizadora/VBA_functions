Attribute VB_Name = "RecebimentosAdiantadosTU"
Function PreencherRecebimentosAdiantadosTU( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' PrintIniFuncao("R. Adiantados TU")
    PreencherRecebimentosAdiantadosTU = ImplementacaoBuscarInfosUnidades(mes_offset, coluna_data, "Recebimentos", 2)

End Function