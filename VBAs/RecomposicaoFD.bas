Attribute VB_Name = "RecomposicaoFD"
Function PreencherRecomposicaoFD( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    PreencherRecomposicaoFD = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "InfosFundos", 4)

End Function