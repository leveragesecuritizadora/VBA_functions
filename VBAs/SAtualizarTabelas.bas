Attribute VB_Name = "SAtualizarTabelas"
Sub AtualizarTabelas()

    Dim conn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim pastaSQL As String
    Dim arquivo As String
    Dim sql As String
    Dim nomeBase As String

    ' iterando sobre os sql
    dim teste As String
    teste = BaixarTexto("https://raw.githubusercontent.com/leveragesecuritizadora/VBA_functions/main/manifest_sql.txt")

    LimparTerminal "SQLLLL"

    Debug.Print teste

    ' ' Conexão
    ' Set conn = CreateObject("ADODB.Connection")
    ' conn.Open _
    '     "Provider=MSOLEDBSQL;" & _
    '     "Server=sqlserver-emissions-prod.database.windows.net,1433;" & _
    '     "Database=sqldb-emissions-dw-prod;" & _
    '     "User ID=app_read;" & _
    '     "Password=JeLBfsQRPt3e5;" & _
    '     "Encrypt=Yes;" & _
    '     "TrustServerCertificate=Yes;"

    ' arquivo = Dir(pastaSQL & "*.sql")

    ' Do While arquivo <> ""

    '     nomeBase = Replace(arquivo, ".sql", "")

    '     ' 1. Planilha
    '     Set ws = GetOrCreateSheet(nomeBase)

    '     ' 1.1 Cor da aba
    '     ws.Tab.Color = RGB(139, 0, 0)

    '     ' 1.2 Ocultando planilhas criadas
    '     ' ws.Visible = xlSheetHidden

    '     ' 2. SQL
    '     sql = LerArquivoTexto(pastaSQL & arquivo)

    '     ' 3. Executa
    '     Set rs = CreateObject("ADODB.Recordset")
    '     rs.Open sql, conn

    '     ' 4. Tabela
    '     Set tbl = GetOrCreateTable(ws, nomeBase)

    '     ' 5. Limpa apenas os dados
    '     If Not tbl.DataBodyRange Is Nothing Then
    '         tbl.DataBodyRange.ClearContents
    '     End If

    '     ' 6. Cabeçalhos vindos do SQL
    '     CabecalhoDaConsulta ws, tbl.Range.Cells(1, 1), rs

    '     ' 7. Dados
    '     tbl.Range.Cells(2, 1).CopyFromRecordset rs

    '     rs.Close
    '     arquivo = Dir

    ' Loop

    ' conn.Close

End Sub