Attribute VB_Name = "PMTSenior"
Public Function PreencherPMTSenior( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    Debug.Print Now() & " ======================PMTSenior"
    PreencherPMTSenior = ImplementacaoBuscarInfosEmissao(mes_offset, coluna_data, "Juros", 7, "senior") 

End Function