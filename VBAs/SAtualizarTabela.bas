Attribute VB_Name = "SAtualizarTabela"
Sub AtualizarTabela()

    Dim conn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim pastaSQL As String
    Dim arquivo As String
    Dim sql As String
    Dim nomeBase As String

    ' Pasta dos .sql
    pastaSQL = Environ("USERPROFILE") & "\OneDrive - Leverage\Área de Trabalho\repos\VBA_functions\consultas\"

    ' Conexão
    Set conn = CreateObject("ADODB.Connection")
    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Server=sqlserver-emissions-prod.database.windows.net,1433;" & _
        "Database=sqldb-emissions-dw-prod;" & _
        "User ID=app_read;" & _
        "Password=JeLBfsQRPt3e5;" & _
        "Encrypt=Yes;" & _
        "TrustServerCertificate=Yes;"

    arquivo = Dir(pastaSQL & "*.sql")

    Do While arquivo <> ""

        nomeBase = Replace(arquivo, ".sql", "")

        ' 1. Planilha
        Set ws = GetOrCreateSheet(nomeBase)

        ' 2. SQL
        sql = LerArquivoTexto(pastaSQL & arquivo)

        ' 3. Executa
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open sql, conn

        ' 4. Tabela
        Set tbl = GetOrCreateTable(ws, nomeBase)

        ' 5. Limpa dados antigos
        If Not tbl.DataBodyRange Is Nothing Then
            tbl.DataBodyRange.ClearContents
        End If

        ' 6. Preenche
        tbl.Range(2, 1).CopyFromRecordset rs

        rs.Close
        arquivo = Dir

    Loop

    conn.Close

End Sub
