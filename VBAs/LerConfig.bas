Attribute VB_Name = "LerConfig"
Function LerConfig(caminho As String) As Object
    Dim dict As Object
    Dim linha As String, partes() As String
    Dim fso As Object, ts As Object

    Set dict = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(caminho, 1)

    Do Until ts.AtEndOfStream
        linha = Trim(ts.ReadLine)
        If linha <> "" And InStr(linha, "=") > 0 Then
            partes = Split(linha, "=")
            dict(partes(0)) = partes(1)
        End If
    Loop

    ts.Close
    Set LerConfig = dict
End Function
