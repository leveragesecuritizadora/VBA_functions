Attribute VB_Name = "FFormatarDataString"
Function FormatarDataString( _
    data_base As Variant, _ 
    mes_offset As Variant _
) As Variant

    FormatarDataString = Format(DateSerial(Year(data_base), Month(data_base) + mes_offset, 1), "dd/mm/yyyy")

End Function