# Process LSOA breakdown data for PiP areas
$projectPath = "C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer"

Write-Host "Processing LSOA breakdown data..." -ForegroundColor Cyan

# Read lookup file
$lookup = Import-Csv (Join-Path $projectPath "lsoa_msoa_la_region_lookup_csv.csv")
Write-Host "Loaded $($lookup.Count) LSOA-MSOA mappings" -ForegroundColor Green

# Create LSOA to MSOA lookup
$lsoaToMsoa = @{}
foreach ($row in $lookup) {
    $lsoaToMsoa[$row.LSOA21CD] = $row.MSOA21CD
}

# Read LSOA HLNM data - skip header rows (rows 1-5), data starts row 6
$hlnmRaw = Get-Content (Join-Path $projectPath "Hyper-Local-Need-Measure-2025-_csv.csv")
$dataRows = $hlnmRaw | Select-Object -Skip 5

# Parse LSOA data
$lsoaData = @()
foreach ($line in $dataRows) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $cols = $line -split ','
    if ($cols.Count -ge 10 -and $cols[0] -match '^E01') {
        $lsoaData += @{
            lsoa_code = $cols[0]
            lsoa_name = $cols[1]
            la_code = $cols[2]
            la_name = $cols[3]
            score = $cols[4]
            typology = $cols[9].Trim()
            rank = $cols[11].Trim()
        }
    }
}
Write-Host "Parsed $($lsoaData.Count) LSOAs with typology data" -ForegroundColor Green

# Read existing data.json to get PiP MSOA codes
$dataJson = Get-Content (Join-Path $projectPath "data.json") -Raw | ConvertFrom-Json
$pipMsoas = $dataJson.areas | ForEach-Object { $_.msoa_code }
Write-Host "PiP areas: $($pipMsoas.Count)" -ForegroundColor Green

# Aggregate by MSOA
$msoaBreakdown = @{}
foreach ($lsoa in $lsoaData) {
    $msoa = $lsoaToMsoa[$lsoa.lsoa_code]
    if (-not $msoa) { continue }

    if (-not $msoaBreakdown.ContainsKey($msoa)) {
        $msoaBreakdown[$msoa] = @{
            total = 0
            critical = 0
            priority = 0
            support = 0
            lsoas = @()
        }
    }

    $msoaBreakdown[$msoa].total++
    $msoaBreakdown[$msoa].lsoas += @{
        code = $lsoa.lsoa_code
        name = $lsoa.lsoa_name
        typology = $lsoa.typology
        rank = $lsoa.rank
    }

    switch -Wildcard ($lsoa.typology) {
        "*Critical*" { $msoaBreakdown[$msoa].critical++ }
        "*Priority*" { $msoaBreakdown[$msoa].priority++ }
        "*Support*" { $msoaBreakdown[$msoa].support++ }
    }
}

# Update PiP areas with LSOA breakdown
$matchCount = 0
foreach ($area in $dataJson.areas) {
    $msoa = $area.msoa_code
    if ($msoaBreakdown.ContainsKey($msoa)) {
        $breakdown = $msoaBreakdown[$msoa]
        $matchCount++

        $area | Add-Member -NotePropertyName "lsoa_total" -NotePropertyValue $breakdown.total -Force
        $area | Add-Member -NotePropertyName "lsoa_critical" -NotePropertyValue $breakdown.critical -Force
        $area | Add-Member -NotePropertyName "lsoa_priority" -NotePropertyValue $breakdown.priority -Force
        $area | Add-Member -NotePropertyName "lsoa_support" -NotePropertyValue $breakdown.support -Force

        # Store individual LSOA details
        $area | Add-Member -NotePropertyName "lsoas" -NotePropertyValue $breakdown.lsoas -Force

        Write-Host "  $msoa ($($area.neighbourhood_name)): $($breakdown.critical)C / $($breakdown.priority)P / $($breakdown.support)S" -ForegroundColor Gray
    } else {
        Write-Host "  NO LSOA DATA: $msoa ($($area.neighbourhood_name))" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Matched $matchCount / $($dataJson.areas.Count) areas with LSOA data" -ForegroundColor Green

# Update metadata
$dataJson.metadata | Add-Member -NotePropertyName "lsoa_breakdown_added" -NotePropertyValue (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ") -Force

# Save
$dataJson | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $projectPath "data.json") -Encoding UTF8
Write-Host "Saved to data.json" -ForegroundColor Green
