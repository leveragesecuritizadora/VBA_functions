Function PreencherRecebimentosTotais( _
    Optional unidade As String = "Unidade", _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

'    PrintIniFuncao("R. Totais")
   unidade = NormalizarTexto(unidade)
   PreencherRecebimentosTotais = ValorPrimeiroMatch(mes_offset, coluna_data, "Recebimentos", 5, Array(unidade))

End Function