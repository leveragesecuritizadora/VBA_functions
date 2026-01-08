Attribute VB_Name = "SImportarModulos"
Private Sub ImportarModulos()
    Dim pasta As String
    Dim arquivo As String

    ' Debug.Print "Dentro ImportarModulos"

    pasta = Environ("TEMP") & "\vba\"
    arquivo = Dir(pasta & "*.bas")
    ' Debug.Print "meio ImportarModulos"

    Do While arquivo <> ""
        If arquivo <> "OrquestradorAtualizacoesVBAs.bas" Then
            ThisWorkbook.VBProject.VBComponents.Import pasta & arquivo
        End If
        arquivo = Dir
    Loop
    ' Debug.Print "saindo ImportarModulos"

End Sub