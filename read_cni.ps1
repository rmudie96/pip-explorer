$ErrorActionPreference = "SilentlyContinue"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$files = @(
    "MSOA_Community Needs Index 2023_ Community Needs rank_2023-01-01.xlsx",
    "MSOA_Community Needs Index 2023_ Active and Engaged Community rank_2023-01-01.xlsx",
    "MSOA_Community Needs Index 2023_ Civic Assets rank_2023-01-01.xlsx",
    "MSOA_Community Needs Index 2023_ Connectedness rank_2023-01-01.xlsx"
)

$basePath = "C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer"

foreach ($file in $files) {
    $filePath = Join-Path $basePath $file
    Write-Host "=== $file ==="

    $workbook = $excel.Workbooks.Open($filePath)
    $sheet = $workbook.Sheets.Item(1)
    $usedRange = $sheet.UsedRange

    Write-Host "Rows: $($usedRange.Rows.Count), Columns: $($usedRange.Columns.Count)"

    # Get column headers
    $headers = @()
    for ($col = 1; $col -le $usedRange.Columns.Count; $col++) {
        $headers += $sheet.Cells.Item(1, $col).Text
    }
    Write-Host "Columns: $($headers -join ', ')"

    # Sample data
    Write-Host "Sample (row 2):"
    $rowData = @()
    for ($col = 1; $col -le $usedRange.Columns.Count; $col++) {
        $rowData += $sheet.Cells.Item(2, $col).Text
    }
    Write-Host "  $($rowData -join ' | ')"

    $workbook.Close($false)
    Write-Host ""
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
