Attribute VB_Name = "PMTSubordinadoSerie"
Public Function PreencherPMTSubordinadoSerie( _
    n_serie As Integer, _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    Debug.Print("PMTSubordinadoSerie")
    PreencherPMTSubordinadoSerie = SomarValoresMultiplasLinhas(mes_offset, coluna_data, "Juros", 7, Array(n_serie, "subordinado")) 

End Function