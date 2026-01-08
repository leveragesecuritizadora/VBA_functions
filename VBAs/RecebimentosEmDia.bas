Attribute VB_Name = "RecebimentosEmDia"
Function PreencherRecebimentosEmDia( _
    Optional unidade As String = "Unidade", _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant
   ' PrintIniFuncao("R. em dia")
   unidade = NormalizarTexto(unidade)
    PreencherRecebimentosEmDia = ValorPrimeiroMatch(mes_offset, coluna_data, "Recebimentos", 4, Array(unidade))
End Function






