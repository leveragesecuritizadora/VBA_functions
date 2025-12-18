Attribute VB_Name = "FNomeEmissao"
Function NomeEmissao()
    dim emissao as String
    Dim nomePlanilha As String

    nomePlanilha = Application.Caller.Parent.Parent.Name

    emissao = nomePlanilha
    emissao = Replace(emissao, "CRI ", "")
    emissao = Replace(emissao, ".", "")
    emissao = Replace(emissao, "Cascata", "")
    emissao = Replace(emissao, "Automatizada", "")
    emissao = Replace(emissao, "VBA", "")
    emissao = Replace(emissao, "xlsm", "")

    NomeEmissao = Trim(emissao)

End Function