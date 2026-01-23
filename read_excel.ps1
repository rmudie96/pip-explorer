$ErrorActionPreference = "SilentlyContinue"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$filePath = "C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer\HLNM_MSOA.xlsx"
$workbook = $excel.Workbooks.Open($filePath)

Write-Host "=== HLNM_MSOA.xlsx ==="
Write-Host "Number of sheets: $($workbook.Sheets.Count)"

foreach ($sheet in $workbook.Sheets) {
    Write-Host ""
    Write-Host "Sheet: $($sheet.Name)"
    $usedRange = $sheet.UsedRange
    Write-Host "Rows: $($usedRange.Rows.Count), Columns: $($usedRange.Columns.Count)"

    # Get column headers (first row)
    $headers = @()
    for ($col = 1; $col -le [Math]::Min($usedRange.Columns.Count, 50); $col++) {
        $cellValue = $sheet.Cells.Item(1, $col).Text
        if ($cellValue) {
            $headers += $cellValue
        }
    }
    Write-Host "Columns: $($headers -join ', ')"
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
