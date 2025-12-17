$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open("C:\Users\Caique\OneDrive - Leverage\Área de Trabalho\teste.xlsm")

$vbProj = $wb.VBProject

# 1. Importar módulos
$path = "$env:USERPROFILE\OneDrive - Leverage\Área de Trabalho\repos\VBA_functions\VBAs\"

Get-ChildItem "$path\*.bas" | ForEach-Object {
    $vbProj.VBComponents.Import($_.FullName)
}

# 2. Executar subs recém-importadas
# $excel.Run("'teste.xlsm'!AtualizarModulos")
$excel.Run("'teste.xlsm'!SAtualizarTabela")

$wb.Save()
$wb.Close()
$excel.Quit()
