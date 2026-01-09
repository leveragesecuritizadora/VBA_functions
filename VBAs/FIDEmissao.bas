Attribute VB_Name = "FIDEmissao"

Function IDEmissao() As Variant

    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim emissao As String

    ' =========================
    ' 1. Trata o nome do arquivo
    ' =========================
    emissao = LCase(ThisWorkbook.Name)

    emissao = Replace(emissao, "cri ", "")
    emissao = Replace(emissao, "temp", "")
    emissao = Replace(emissao, "_", "")
    emissao = Replace(emissao, " - ", "")
    emissao = Replace(emissao, ".", "")
    emissao = Replace(emissao, "cascata", "")
    emissao = Replace(emissao, "automatizada", "")
    emissao = Replace(emissao, "vba", "")
    emissao = Replace(emissao, "xlsm", "")
    emissao = Replace(emissao, "cra", "")
    emissao = Replace(emissao, "(", "")
    emissao = Replace(emissao, ")", "")
    emissao = Trim(Replace(emissao, "xlsx", ""))

    LimparTerminal "Buscando id da emissão: " & emissao

    ' =========================
    ' 2. Monta SQL
    ' =========================
    sql = "SELECT DW.getEmissaoId('" & emissao & "') AS emissao_id;"

    ' =========================
    ' 3. Conecta no BD
    ' =========================
    Set conn = CreateObject("ADODB.Connection")
    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Server=sqlserver-emissions-prod.database.windows.net,1433;" & _
        "Database=sqldb-emissions-dw-prod;" & _
        "User ID=app_read;" & _
        "Password=JeLBfsQRPt3e5;" & _
        "Encrypt=Yes;" & _
        "TrustServerCertificate=Yes;"

    ' =========================
    ' 4. Executa
    ' =========================
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn

    ' =========================
    ' 5. Retorna o valor
    ' =========================
    If (Not rs.EOF) And (Not IsNull(rs.Fields(0).Value)) Then
        IDEmissao = CLng(rs.Fields(0).Value)
    Else
        MsgBox Now & " Emissão " & emissao & " não encontrada no servidor"
        IDEmissao = False
    End If

    ' =========================
    ' 6. Limpa
    ' =========================
    rs.Close
    conn.Close

    Set rs = Nothing
    Set conn = Nothing

End Function