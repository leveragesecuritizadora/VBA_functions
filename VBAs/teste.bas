Attribute VB_Name = "teste"
Function PreencheRecebimento( _
    Optional unidade As String = "Unidade", _
    Optional tipo_recebimento As String = "total", _
    Optional dado_historico As Variant, _
    Optional mes_desejado As Variant = False, _
    Optional mes_offset As Integer = -1, _
    Optional place_holder As Variant = "-", _
    Optional nome_fonte As String = "Recebimentos" _
) As Variant

    Dim wb As Workbook
    Set wb = Workbooks.Open("C:\Users\Caique\OneDrive - Leverage\Área de Trabalho\repos\VBA_functions\planilhas\Dados_Emissoes.xlsx")

    Dim ws As Worksheet
    Set ws = wb.Sheets("Recebimentos")

    MsgBox ws.Range("A1").Value
    
    Exit Function

' --- [9] Tratamento gen_rico de erro inesperado ---
ErroHandler:
    PreencheRecebimento = "--"
End Function



