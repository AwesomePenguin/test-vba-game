# ImportVBA.ps1
$excel = New-Object -ComObject Excel.Application

# Use Resolve-Path to get the absolute path from the relative path
$workbookPath = Resolve-Path "./test-vba-game.xlsm"
$workbook = $excel.Workbooks.Open($workbookPath)

$importPath = Resolve-Path "./code"
Get-ChildItem -Path $importPath -Filter *.bas | ForEach-Object {
    $workbook.VBProject.VBComponents.Import($_.FullName)
}

$workbook.Save()
$workbook.Close()
$excel.Quit()
Pause