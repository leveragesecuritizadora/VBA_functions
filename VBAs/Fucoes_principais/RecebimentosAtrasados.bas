Attribute VB_Name = "RecebimentosAtrasados"
Function PreencherRecebimentosAtrasados( _
    Optional unidade As String = "Unidade", _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    

   ' PrintIniFuncao("R. Atrasados")
    PreencherRecebimentosAtrasados = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "Recebimentos", 3, unidade)

End Function






