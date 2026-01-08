Attribute VB_Name = "RecebimentosAtrasados"
Function PreencherRecebimentosAtrasados( _
    Optional unidade As String = "Unidade", _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    

   ' Debug.Print("R. Atrasados")
   unidade = NormalizarTexto(unidade)
    PreencherRecebimentosAtrasados = ValorPrimeiroMatch(mes_offset, coluna_data, "Recebimentos", 3, Array(unidade))

End Function






