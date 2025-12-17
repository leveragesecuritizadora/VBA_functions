$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open("C:\Users\Caique\OneDrive - Leverage\Área de Trabalho\teste.xlsm")

$excel.Run("'teste.xlsm'!AtualizarModulos")
$excel.Run("'teste.xlsm'!AtualizarTabela_SQLServer")

$wb.Save()
$wb.Close()
$excel.Quit()
