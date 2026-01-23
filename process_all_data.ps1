# Fast combined processing script
$projectPath = "C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer"
$dataJson = Get-Content (Join-Path $projectPath "data.json") -Raw | ConvertFrom-Json
$pipMsoas = @{}
$dataJson.areas | ForEach-Object { $pipMsoas[$_.msoa_code] = $_ }
Write-Host "PiP MSOAs: $($pipMsoas.Count)" -ForegroundColor Cyan

# === CNI DATA ===
Write-Host "Processing CNI data..." -ForegroundColor Cyan
$cniFiles = @{
    "civic_assets_rank" = "MSOA_Community Needs Index 2023_ Civic Assets rank_2023-01-01_csv.csv"
    "connectedness_rank" = "MSOA_Community Needs Index 2023_ Connectedness rank_2023-01-01_csv.csv"
    "active_engaged_rank" = "MSOA_Community Needs Index 2023_ Active and Engaged Community rank_2023-01-01_csv.csv"
}

foreach ($key in $cniFiles.Keys) {
    $file = Join-Path $projectPath $cniFiles[$key]
    if (Test-Path $file) {
        $data = Import-Csv $file
        $matched = 0
        foreach ($row in $data) {
            if ($pipMsoas.ContainsKey($row.'Area Code')) {
                $pipMsoas[$row.'Area Code'] | Add-Member -NotePropertyName $key -NotePropertyValue ([int]$row.Value) -Force
                $matched++
            }
        }
        Write-Host "  $key : matched $matched" -ForegroundColor Green
    } else {
        Write-Host "  $key : FILE NOT FOUND" -ForegroundColor Yellow
    }
}

# === LSOA BREAKDOWN ===
Write-Host "Processing LSOA breakdown..." -ForegroundColor Cyan

# Build LSOA->MSOA lookup (only for PiP MSOAs to save memory)
$lookup = Import-Csv (Join-Path $projectPath "lsoa_msoa_la_region_lookup_csv.csv")
$pipLsoas = @{}
foreach ($row in $lookup) {
    if ($pipMsoas.ContainsKey($row.MSOA21CD)) {
        $pipLsoas[$row.LSOA21CD] = $row.MSOA21CD
    }
}
Write-Host "  LSOAs in PiP areas: $($pipLsoas.Count)" -ForegroundColor Green

# Initialize counters
foreach ($msoa in $pipMsoas.Keys) {
    $pipMsoas[$msoa] | Add-Member -NotePropertyName "lsoa_total" -NotePropertyValue 0 -Force
    $pipMsoas[$msoa] | Add-Member -NotePropertyName "lsoa_critical" -NotePropertyValue 0 -Force
    $pipMsoas[$msoa] | Add-Member -NotePropertyName "lsoa_priority" -NotePropertyValue 0 -Force
    $pipMsoas[$msoa] | Add-Member -NotePropertyName "lsoa_support" -NotePropertyValue 0 -Force
}

# Parse LSOA HLNM - only rows matching our LSOAs
$hlnmFile = Join-Path $projectPath "Hyper-Local-Need-Measure-2025-_csv.csv"
$reader = [System.IO.StreamReader]::new($hlnmFile)
$lineNum = 0
while ($null -ne ($line = $reader.ReadLine())) {
    $lineNum++
    if ($lineNum -le 5) { continue } # Skip headers
    $cols = $line -split ','
    if ($cols.Count -lt 10) { continue }
    $lsoaCode = $cols[0]
    if (-not $pipLsoas.ContainsKey($lsoaCode)) { continue }

    $msoa = $pipLsoas[$lsoaCode]
    $typology = $cols[9]

    $pipMsoas[$msoa].lsoa_total++
    if ($typology -match "Critical") { $pipMsoas[$msoa].lsoa_critical++ }
    elseif ($typology -match "Priority") { $pipMsoas[$msoa].lsoa_priority++ }
    elseif ($typology -match "Support") { $pipMsoas[$msoa].lsoa_support++ }
}
$reader.Close()
Write-Host "  LSOA parsing complete" -ForegroundColor Green

# Summary
foreach ($area in $dataJson.areas) {
    $c = $area.lsoa_critical; $p = $area.lsoa_priority; $s = $area.lsoa_support
    Write-Host "  $($area.neighbourhood_name): ${c}C/${p}P/${s}S" -ForegroundColor Gray
}

# Save
$dataJson | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $projectPath "data.json") -Encoding UTF8
Write-Host "Saved data.json" -ForegroundColor Green
