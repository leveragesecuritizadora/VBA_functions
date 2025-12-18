Attribute VB_Name = "RecebimentosTotais"
Function PreencherRecebimentosTotais( _
    Optional unidade As String = "Unidade", _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' PrintIniFuncao("R. Totais")
    PreencherRecebimentosTotais = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "Recebimentos", 5, unidade)

End Function