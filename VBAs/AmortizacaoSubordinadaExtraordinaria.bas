Attribute VB_Name = "AMEXSubordinada"
Public Function PreencherAmexSubordinada( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' PrintIniFuncao("AMEXSub")
    PreencherAmexSubordinada = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "Juros", 6, "subordinada")

End Function