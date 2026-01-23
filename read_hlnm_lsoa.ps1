$ErrorActionPreference = "SilentlyContinue"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$filePath = "C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer\Hyper-Local-Need-Measure-2025.xlsx"
$workbook = $excel.Workbooks.Open($filePath)

Write-Host "=== Hyper-Local-Need-Measure-2025.xlsx ==="
Write-Host "Number of sheets: $($workbook.Sheets.Count)"

foreach ($sheet in $workbook.Sheets) {
    Write-Host ""
    Write-Host "Sheet: $($sheet.Name)"
    $usedRange = $sheet.UsedRange
    Write-Host "Rows: $($usedRange.Rows.Count), Columns: $($usedRange.Columns.Count)"

    # Get column headers
    $headers = @()
    for ($col = 1; $col -le [Math]::Min($usedRange.Columns.Count, 30); $col++) {
        $cellValue = $sheet.Cells.Item(1, $col).Text
        if ($cellValue) {
            $headers += $cellValue
        }
    }
    Write-Host "Columns: $($headers -join ', ')"

    # Sample first 3 data rows
    Write-Host "Sample data (rows 2-4):"
    for ($row = 2; $row -le [Math]::Min(4, $usedRange.Rows.Count); $row++) {
        $rowData = @()
        for ($col = 1; $col -le [Math]::Min(8, $usedRange.Columns.Count); $col++) {
            $rowData += $sheet.Cells.Item($row, $col).Text
        }
        Write-Host "  $($rowData -join ' | ')"
    }
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
