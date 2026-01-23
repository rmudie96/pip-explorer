$ErrorActionPreference = "SilentlyContinue"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$file = "MSOA_Community Needs Index 2023_ Community Needs rank_2023-01-01.xlsx"
$basePath = "C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer"
$filePath = Join-Path $basePath $file

Write-Host "=== $file ==="
$workbook = $excel.Workbooks.Open($filePath)

Write-Host "Number of sheets: $($workbook.Sheets.Count)"

foreach ($sheet in $workbook.Sheets) {
    Write-Host ""
    Write-Host "Sheet: $($sheet.Name)"
    $usedRange = $sheet.UsedRange
    Write-Host "Rows: $($usedRange.Rows.Count), Columns: $($usedRange.Columns.Count)"

    # Get all column headers
    Write-Host "Columns:"
    for ($col = 1; $col -le $usedRange.Columns.Count; $col++) {
        $cellValue = $sheet.Cells.Item(1, $col).Text
        Write-Host "  $col : $cellValue"
    }

    # First 10 rows
    Write-Host "First 10 rows:"
    for ($row = 1; $row -le [Math]::Min(10, $usedRange.Rows.Count); $row++) {
        $rowData = @()
        for ($col = 1; $col -le [Math]::Min(6, $usedRange.Columns.Count); $col++) {
            $rowData += $sheet.Cells.Item($row, $col).Text
        }
        Write-Host "  Row $row : $($rowData -join ' | ')"
    }
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
