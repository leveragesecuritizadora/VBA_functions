VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If _
        Sh.Name = "Recebimentos" Or _
        Sh.Name = "Juros" Or _
        Sh.Name = "SaldoResidual" Or _
        Sh.Name = "SalFR" Or _
        Sh.Name = "VP" Or _
        Sh.Name = "PrevistosDevidos" Or _
        Sh.Name = "SalminFR" Or _
        Sh.Name = "SaldoFD" _
    Then
        Application.CalculateFull
    End If
End Sub

