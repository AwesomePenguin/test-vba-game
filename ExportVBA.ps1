# ExportVBA.ps1
$excel = New-Object -ComObject Excel.Application

# Use Resolve-Path to get the absolute path from the relative path
try {
    $workbookPath = Resolve-Path "./test-vba-game.xlsm"
    Write-Output "Resolved workbook path: $workbookPath"
} catch {
    Write-Error "Failed to resolve workbook path. Ensure the file exists."
    exit 1
}

# Open the workbook
try {
    $workbook = $excel.Workbooks.Open($workbookPath)
    if ($null -eq $workbook) {
        throw "Workbook could not be opened."
    }
} catch {
    Write-Error "Failed to open workbook: $_"
    $excel.Quit()
    exit 1
}

# Use Resolve-Path to get the absolute path for the export directory
try {
    $exportPath = Resolve-Path "./code"
    Write-Output "Resolved export path: $exportPath"
} catch {
    Write-Error "Failed to resolve export path. Ensure the directory exists."
    $workbook.Close($false)
    $excel.Quit()
    exit 1
}

# Export VBA components
try {
    $workbook.VBProject.VBComponents | ForEach-Object {
        $component = $_
        if ($null -ne $component) {
            $fileName = Join-Path $exportPath ($component.Name + ".bas")
            try {
                $component.Export($fileName)
                Write-Output "Exported $($component.Name) to $fileName"
            } catch {
                Write-Error "Failed to export component $($component.Name): $_"
            }
        } else {
            Write-Output "Skipped a null component."
        }
    }
} catch {
    Write-Error "Failed to export VBA components: $_"
}

# Close the workbook and quit Excel
$workbook.Close($false)
$excel.Quit()