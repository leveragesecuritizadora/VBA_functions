Attribute VB_Name = "RecebimentosAdiantadosTU"
Function PreencherRecebimentosAdiantadosTU( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

'    Debug.Print("R. Adiantados TU")
   unidade = NormalizarTexto(unidade)
    PreencherRecebimentosAdiantadosTU = SomarValoresMultiplasLinhas(mes_offset, coluna_data, "Recebimentos", 2, Array("*"))

End Function