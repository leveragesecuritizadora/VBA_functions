Attribute VB_Name = "FNomeEmissao"
Function NomeEmissao()
    dim emissao as String
    Dim nomePlanilha As String

    nomePlanilha = Application.Caller.Parent.Parent.Name

    emissao = nomePlanilha
    ' MeuPrint nomePlanilha
    emissao = Replace(emissao, "CRI ", "")
    emissao = Replace(emissao, " - ", "")
    ' MeuPrint "rm CRI ", emissao
    emissao = Replace(emissao, ".", "")
    ' MeuPrint "rm .", emissao
    emissao = Replace(emissao, "Cascata", "")
    ' MeuPrint "rm Cascata", emissao
    emissao = Replace(emissao, "Automatizada", "")
    ' MeuPrint "rm Automatizada", emissao
    emissao = Replace(emissao, "VBA", "")
    ' MeuPrint "rm VBA", emissao
    emissao = Replace(emissao, "xlsm", "")
    ' MeuPrint "rm xlsm", emissao

    NomeEmissao = Trim(emissao)

End Function