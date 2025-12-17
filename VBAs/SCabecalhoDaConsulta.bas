Attribute VB_Name = "SCabecalhoDaConsulta"
Public Sub CabecalhoDaConsulta(ws As Worksheet, _
                                     startCell As Range, _
                                     rs As Object)
    Dim i As Long

    For i = 0 To rs.Fields.Count - 1
        startCell.Offset(0, i).Value = rs.Fields(i).Name
    Next i
End Sub
