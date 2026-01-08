Attribute VB_Name = "RecebimentosAtrasadosTU"
Function PreencherRecebimentosAtrasadosTU( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' Debug.Print("R. Atrasados TU")
   unidade = NormalizarTexto(unidade)
    PreencherRecebimentosAtrasadosTU = SomarValoresMultiplasLinhas(mes_offset, coluna_data, "Recebimentos", 3, Array("*"))

End Function