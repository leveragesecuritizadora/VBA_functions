Attribute VB_Name = "FSanitizarXLSM"
Sub SanitizarXLSM()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim caminho As String
    
    Set wb = ThisWorkbook
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Garante que tudo esteja calculado
    wb.RefreshAll
    Application.CalculateFull
    
    ' Converte fórmulas em valores
    For Each ws In wb.Worksheets
        With ws.UsedRange
            .Value = .Value
        End With
    Next ws
    
    Application.Calculation = xlCalculationAutomatic
    
    ' Caminho do novo arquivo
    caminho = Replace(wb.FullName, ".xlsm", "_sanitizado.xlsx")
    
    ' Salva como XLSX (sem macros)
    wb.SaveAs Filename:=caminho, FileFormat:=xlOpenXMLWorkbook
    
    Application.ScreenUpdating = True
    
    MsgBox "Arquivo sanitizado com sucesso!" & vbCrLf & caminho, vbInformation
End Sub
