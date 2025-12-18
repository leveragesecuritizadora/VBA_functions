Attribute VB_Name = "SSanitizarXLSM"
Sub SanitizarXLSM()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim ole As OLEObject
    Dim caminho As String

    Set wb = ThisWorkbook

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' ==========================
    ' 1) Garante cálculo
    ' ==========================
    wb.RefreshAll
    Application.CalculateFull

    ' ==========================
    ' 2) Converte fórmulas em valores
    ' ==========================
    For Each ws In wb.Worksheets
        ws.UsedRange.Value = ws.UsedRange.Value
    Next ws

    Application.Calculation = xlCalculationAutomatic

    ' ==========================
    ' 3) Remove botões
    ' ==========================
    For Each ws In wb.Worksheets

        ' Remove botões Form Control
        For Each shp In ws.Shapes
            If shp.Type = msoFormControl Then
                shp.Delete
            End If
        Next shp

        ' Remove botões ActiveX
        For Each ole In ws.OLEObjects
            ole.Delete
        Next ole

    Next ws

    ' ==========================
    ' 4) Salva como XLSX
    ' ==========================
    caminho = Replace(wb.FullName, ".xlsm", "_sanitizado.xlsx")
    wb.SaveAs Filename:=caminho, FileFormat:=xlOpenXMLWorkbook

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Arquivo sanitizado com sucesso!" & vbCrLf & caminho, vbInformation

End Sub
