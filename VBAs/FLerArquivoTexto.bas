Attribute VB_Name = "FLerArquivoTexto"
Function LerArquivoTexto(caminho As String) As String
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(caminho, 1) ' ForReading
    LerArquivoTexto = ts.ReadAll
    ts.Close
End Function
