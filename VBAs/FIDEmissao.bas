Attribute VB_Name = "FIDEmissao"
Function IDEmissao() As Integer
    Dim id As Integer
    Dim emissao As String

    Debug.Print emissao

    emissao = LCase(ThisWorkbook.Name)
    ' MeuPrint nomePlanilha
    emissao = Replace(emissao, "cri ", "")
    emissao = Replace(emissao, "temp", "")
    emissao = Replace(emissao, "_", "")
    emissao = Replace(emissao, " - ", "")
    emissao = Replace(emissao, ".", "")
    emissao = Replace(emissao, "cascata", "")
    emissao = Replace(emissao, "automatizada", "")
    emissao = Replace(emissao, "vba", "")
    emissao = Replace(emissao, "xlsm", "")
    emissao = Trim(Replace(emissao, "xlsx", ""))

    Debug.Print emissao

End Function