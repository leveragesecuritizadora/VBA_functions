Attribute VB_Name = "Recebimentos"
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
    Set wb = Workbooks.Open("https://leveragesec-my.sharepoint.com/personal/caique_leveragesec_com_br/Documents/Dados_Emissoes.xlsx?web=1")

    Dim ws As Worksheet
    Set ws = wb.Sheets("Dados")

    MsgBox ws.Range("A1").Value
    
    Exit Function

' --- [9] Tratamento gen_rico de erro inesperado ---
ErroHandler:
    PreencheRecebimento = "--"
End Function



