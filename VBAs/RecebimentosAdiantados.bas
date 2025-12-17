Attribute VB_Name = "RecebimentosAdiantados"
Function PreencherRecebimentosAdiantados( _
    Optional unidade As String = "Unidade", _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    PrintIniFuncao("R. Adiantado")
    PreencherRecebimentosAdiantados = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "Recebimentos", 2, unidade)

End Function






