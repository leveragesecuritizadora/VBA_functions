Attribute VB_Name = "FIDEmissao"
Function IDEmissao() As Integer
    Dim id As String
    Dim partes() As String

    partes = Split(ThisWorkbook.Name, ".")
    partes = Split(partes(0), " ")

    id = CInt(Split(LCase(partes(0)), "e")(0))

    IDEmissao = id



End Function