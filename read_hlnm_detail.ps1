$ErrorActionPreference = "SilentlyContinue"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$filePath = "C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer\Hyper-Local-Need-Measure-2025.xlsx"
$workbook = $excel.Workbooks.Open($filePath)
$sheet = $workbook.Sheets.Item(1)
$usedRange = $sheet.UsedRange

Write-Host "=== Hyper-Local-Need-Measure-2025.xlsx Detail ==="
Write-Host "Total Rows: $($usedRange.Rows.Count)"
Write-Host "Total Columns: $($usedRange.Columns.Count)"
Write-Host ""

# Get ALL column headers
Write-Host "All Columns:"
for ($col = 1; $col -le $usedRange.Columns.Count; $col++) {
    $cellValue = $sheet.Cells.Item(1, $col).Text
    Write-Host "  Col $col : $cellValue"
}

Write-Host ""
Write-Host "First 15 rows (all columns):"
for ($row = 1; $row -le 15; $row++) {
    $rowData = @()
    for ($col = 1; $col -le [Math]::Min(10, $usedRange.Columns.Count); $col++) {
        $rowData += $sheet.Cells.Item($row, $col).Text
    }
    Write-Host "Row $row : $($rowData -join ' | ')"
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
