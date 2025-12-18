Attribute VB_Name = "CotasSubordinada"
Function PreencherCotasSubordinada( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

   ' Debug.Print "Cotas Subordinada"
    PreencherCotasSubordinada = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "Juros", 2, "subordinada")

End Function