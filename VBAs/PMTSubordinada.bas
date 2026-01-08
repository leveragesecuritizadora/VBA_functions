Attribute VB_Name = "PMTSubordinada"
Public Function PreencherPMTSubordinada( _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    ' Debug.Print Now() & " ======================PMTSub"
    PreencherPMTSubordinada = SomarValoresMultiplasLinhas(mes_offset, coluna_data, "Juros", 7, Array("*", "subordinada")) 

End Function






