Attribute VB_Name = "SApagarModulos"
Private Sub ApagarModulos()
    Dim i As Long
    Dim vbComp As Object

    LimparTerminal "Apagando Módulos Antigos"

    For i = ThisWorkbook.VBProject.VBComponents.Count To 1 Step -1
        Set vbComp = ThisWorkbook.VBProject.VBComponents(i)

        If vbComp.Type = 1 _
           And vbComp.Name <> "Bootloader" _
           And vbComp.Name <> "OrquestradorAtualizacoesVBAs" Then
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
        End If
    Next i
End Sub