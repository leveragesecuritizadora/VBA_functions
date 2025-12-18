Attribute VB_Name = "CotasSenior"
Function PreencherCotasSenior( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    Debug.Print "Cotas Senior"
    PreencherCotasSenior = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "Juros", 2, "senior")

End Function