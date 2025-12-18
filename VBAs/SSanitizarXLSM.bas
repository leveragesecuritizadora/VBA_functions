Attribute VB_Name = "SSanitizarXLSM"

Sub SanitizarXLSM()

    Dim wbOrigem As Workbook
    Dim wbTemp As Workbook
    Dim wbXlsx As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim ole As OLEObject

    Dim caminhoTemp As String
    Dim caminhoXlsx As String
    Dim pastaConsultas As String
    Dim arquivo As String
    Dim nomeAba As String
    Dim i As Long

    Set wbOrigem = ThisWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' ==========================
    ' 1) Cria cópia temporária
    ' ==========================
    caminhoTemp = Replace(wbOrigem.FullName, ".xlsm", "_TEMP.xlsm")
    wbOrigem.SaveCopyAs caminhoTemp

    ' Abre a cópia
    Set wbTemp = Workbooks.Open(caminhoTemp)

    ' ==========================
    ' 2) Garante cálculo
    ' ==========================
    wbTemp.RefreshAll
    Application.CalculateFull

    ' ==========================
    ' 3) Converte fórmulas em valores
    ' ==========================
    For Each ws In wbTemp.Worksheets
        ws.UsedRange.Value = ws.UsedRange.Value
    Next ws

    Application.Calculation = xlCalculationAutomatic

    ' ==========================
    ' 4) Remove botões e OLE
    ' ==========================
    For Each ws In wbTemp.Worksheets
        For Each shp In ws.Shapes
            If shp.Type = msoFormControl Then shp.Delete
        Next shp

        For Each ole In ws.OLEObjects
            ole.Delete
        Next ole
    Next ws

    ' ==========================
    ' 5) Monta caminho do XLSX final
    ' ==========================
    caminhoXlsx = wbOrigem.FullName
    caminhoXlsx = Replace(caminhoXlsx, "CRI ", "")
    caminhoXlsx = Replace(caminhoXlsx, ".", "")
    caminhoXlsx = Replace(caminhoXlsx, "Cascata", "")
    caminhoXlsx = Replace(caminhoXlsx, "Automatizada", "")
    caminhoXlsx = Replace(caminhoXlsx, "VBA", "")
    caminhoXlsx = Replace(caminhoXlsx, "xlsm", _
        " - Cascata " & Format(Date, "mm-yyyy") & ".xlsx")

    ' ==========================
    ' 6) Salva como XLSX (remove VBA)
    ' ==========================
    wbTemp.SaveAs _
        Filename:=caminhoXlsx, _
        FileFormat:=xlOpenXMLWorkbook

    ' O workbook ativo agora é o XLSX
    Set wbXlsx = ActiveWorkbook

    ' Fecha a cópia temporária SEM salvar
    wbTemp.Close SaveChanges:=False

    ' ==========================
    ' 7) Apaga abas baseadas nos .sql
    ' ==========================
    pastaConsultas = Environ("USERPROFILE") & _
        "\OneDrive - Leverage\Área de Trabalho\repos\VBA_functions\consultas\"

    arquivo = Dir(pastaConsultas & "*.sql")

    Do While arquivo <> ""
        nomeAba = Left(arquivo, Len(arquivo) - 4)

        For i = wbXlsx.Worksheets.Count To 1 Step -1
            If wbXlsx.Worksheets(i).Name = nomeAba Then
                wbXlsx.Worksheets(i).Delete
            End If
        Next i

        arquivo = Dir
    Loop

    ' ==========================
    ' 8) Finalização
    ' ==========================
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Arquivo sanitizado gerado com sucesso!" & _
           vbCrLf & caminhoXlsx, vbInformation

End Sub
