Attribute VB_Name = "AtualizarModulos"
Sub AtualizarModulos()

    Dim caminhoPasta As String
    Dim arquivo As String
    Dim vbComp As Object
    
    ' Caminho da sua pasta de módulos (repositório local)
    caminhoPasta = "C:\Users\Caique\OneDrive - Leverage\Área de Trabalho\repos\VBA_functions\VBAs\"

    ' ================================
    ' 1. REMOVER módulos antigos (.bas)
    ' ================================
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 1 Then ' 1 = vbext_ct_StdModule (.bas)
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
        End If
    Next vbComp


    ' ================================
    ' 2. IMPORTAR todos os .bas
    ' ================================
    arquivo = Dir(caminhoPasta & "*.bas")

    Do While arquivo <> ""
        ThisWorkbook.VBProject.VBComponents.Import caminhoPasta & arquivo
        arquivo = Dir
    Loop

    MsgBox "Módulos atualizados com sucesso a partir do repositório!", vbInformation

End Sub
