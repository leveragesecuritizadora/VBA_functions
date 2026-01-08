Attribute VB_Name = "PMTSeniorSerie"
Public Function PreencherPMTSeniorSerie( _
    n_serie As Integer, _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    Debug.Print("PMTSeniorSerie")
    PreencherPMTSeniorSerie = SomarValoresMultiplasLinhas(mes_offset, coluna_data, "Juros", 7, Array(n_serie, "senior")) 

End Function