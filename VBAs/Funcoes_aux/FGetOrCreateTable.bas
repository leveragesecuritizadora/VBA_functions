Attribute VB_Name = "FGetOrCreateTable"
Public Function GetOrCreateTable(ws As Worksheet, nomeTabela As String) As ListObject
    Dim tbl As ListObject

    On Error Resume Next
    Set tbl = ws.ListObjects(nomeTabela)
    On Error GoTo 0

    If tbl Is Nothing Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1"), , xlYes)
        tbl.Name = nomeTabela
    End If

    Set GetOrCreateTable = tbl
End Function
