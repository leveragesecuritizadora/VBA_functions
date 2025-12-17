Attribute VB_Name = "SApagarModulos"
Sub ApagarModulos()

    Dim caminhoPasta As String
    Dim arquivo As String
    Dim vbComp As Object
    
    ' Caminho da sua pasta de módulos (repositório local)
    caminhoPasta = Environ("USERPROFILE") & "\OneDrive - Leverage\Área de Trabalho\repos\VBA_functions\VBAs\"

    ' ================================
    ' 1. REMOVER módulos antigos (.bas)
    ' ================================
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 1 Then ' 1 = vbext_ct_StdModule (.bas)
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
        End If
    Next vbComp
End Sub
