Attribute VB_Name = "SSanitizarXLSM"
Sub SanitizarXLSM()

    Dim wbOrigem As Workbook
    Dim wbXlsx As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim ole As OLEObject
    Dim caminhoXlsx As String
    Dim pastaConsultas As String
    Dim arquivo As String
    Dim nomeAba As String
    Dim i As Long
    Dim contador As Integer

    Set wbOrigem = ThisWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' ==========================
    ' 1) Garante cálculo
    ' ==========================
    wbOrigem.RefreshAll
    Application.CalculateFull

    ' ==========================
    ' 2) Converte fórmulas em valores
    ' ==========================
    For Each ws In wbOrigem.Worksheets
        ws.UsedRange.Value = ws.UsedRange.Value
    Next ws

    Application.Calculation = xlCalculationAutomatic

    ' ==========================
    ' 3) Remove botões
    ' ==========================
    For Each ws In wbOrigem.Worksheets

        For Each shp In ws.Shapes
            If shp.Type = msoFormControl Then shp.Delete
        Next shp

        For Each ole In ws.OLEObjects
            ole.Delete
        Next ole

    Next ws

    ' ==========================
    ' 4) Salva como XLSX
    ' ==========================
    caminhoXlsx = Replace(wbOrigem.FullName, ".xlsm", "_sanitizado.xlsx")
    wbOrigem.SaveAs Filename:=caminhoXlsx, FileFormat:=xlOpenXMLWorkbook

    ' ==========================
    ' 5) Abre o XLSX limpo
    ' ==========================
    Set wbXlsx = Workbooks.Open(caminhoXlsx)

    ' ==========================
    ' 6) Apaga abas com nome dos .sql
    ' ==========================
    pastaConsultas = Environ("USERPROFILE") & "\OneDrive - Leverage\Área de Trabalho\repos\VBA_functions\consultas\"
    arquivo = Dir(pastaConsultas & "*.sql")

    contador = 0
    Do While arquivo <> ""
        nomeAba = Left(arquivo, Len(arquivo) - 4)
        contador = contador + 1
        Debug.print Cstr(contador) & " - " & nomeAba

        For i = wbXlsx.Worksheets.Count To 1 Step -1
            If wbXlsx.Worksheets(i).Name = nomeAba Then
                wbXlsx.Worksheets(i).Delete
            End If
        Next i

        arquivo = Dir
    Loop

    ' ==========================
    ' 7) Salva e fecha o XLSX
    ' ==========================
    wbXlsx.Save
    wbXlsx.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Arquivo sanitizado gerado com sucesso!" & vbCrLf & caminhoXlsx, vbInformation

End Sub
