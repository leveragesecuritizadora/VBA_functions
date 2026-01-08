Attribute VB_Name = "RecomposicaoFR"
Function PreencherRecomposicaoFR( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    PreencherRecomposicaoFR = ValorPrimeiroMatch(mes_offset, coluna_data, "InfosFundos", 7)

End Function