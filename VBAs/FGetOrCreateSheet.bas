Attribute VB_Name = "FGetOrCreateSheet"
Public Function GetOrCreateSheet(nome As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nome)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = nome
    End If

    Set GetOrCreateSheet = ws
End Function
