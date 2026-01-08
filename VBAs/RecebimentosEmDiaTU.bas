Attribute VB_Name = "RecebimentosEmDiaTU"
Function PreencherRecebimentosEmDiaTU(  _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' PrintIniFuncao("R. Em dia TU")
    PreencherRecebimentosEmDiaTU = SomarValoresMultiplasLinhas(mes_offset, coluna_data, "Recebimentos", 4)

End Function