Attribute VB_Name = "FNormalizarTexto"
Function NormalizarTexto(ByVal texto As String) As String
    Dim comAcento As String
    Dim semAcento As String
    Dim i As Long

     comAcento = "áàâãäéèêëíìîïóòôõöúùûüçÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇ"
    semAcento = "aaaaaeeeeiiiiooooouuuucAAAAAEEEEIIIIOOOOOUUUUC"

    ' Remove acentos
    For i = 1 To Len(comAcento)
        texto = Replace(texto, Mid(comAcento, i, 1), Mid(semAcento, i, 1))
    Next i

    ' Converte para minÃºsculo
    NormalizarTexto = LCase(texto)
End Function
