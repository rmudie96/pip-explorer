$ErrorActionPreference = "Stop"
$basePath = "C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer"

Write-Host "=== Pride in Place MSOA Data Extraction ==="

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# ============================================
# 1. Read existing data.json to get PiP MSOAs
# ============================================
Write-Host "1. Reading existing data.json..."
$existingData = Get-Content (Join-Path $basePath "data.json") -Raw | ConvertFrom-Json
$msoaData = @{}
foreach ($area in $existingData.areas) {
    $msoaData[$area.msoa_code] = @{}
    foreach ($prop in $area.PSObject.Properties) {
        $msoaData[$area.msoa_code][$prop.Name] = $prop.Value
    }
}
Write-Host "  Found $($msoaData.Count) Pride in Place areas"

# ============================================
# 2. Read HLNM_MSOA.xlsx (master sheet)
# ============================================
Write-Host "2. Reading HLNM_MSOA.xlsx..."
$hlnmPath = Join-Path $basePath "HLNM_MSOA.xlsx"
$hlnmWorkbook = $excel.Workbooks.Open($hlnmPath)
$masterSheet = $hlnmWorkbook.Sheets.Item("master")
$usedRange = $masterSheet.UsedRange

$allHlnm = @{}
for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
    $areaCode = $masterSheet.Cells.Item($row, 1).Text
    if ($areaCode -match "^E02") {
        $allHlnm[$areaCode] = @{
            growth = [double]$masterSheet.Cells.Item($row, 5).Value2
            energy = [double]$masterSheet.Cells.Item($row, 6).Value2
            crime = [double]$masterSheet.Cells.Item($row, 7).Value2
            opportunity = [double]$masterSheet.Cells.Item($row, 8).Value2
            health = [double]$masterSheet.Cells.Item($row, 9).Value2
        }
        if ($msoaData.ContainsKey($areaCode)) {
            $msoaData[$areaCode]["hlnm_growth"] = $allHlnm[$areaCode].growth
            $msoaData[$areaCode]["hlnm_energy"] = $allHlnm[$areaCode].energy
            $msoaData[$areaCode]["hlnm_crime"] = $allHlnm[$areaCode].crime
            $msoaData[$areaCode]["hlnm_opportunity"] = $allHlnm[$areaCode].opportunity
            $msoaData[$areaCode]["hlnm_health"] = $allHlnm[$areaCode].health
        }
    }
}
Write-Host "  Read $($allHlnm.Count) MSOA HLNM records"
$hlnmWorkbook.Close($false)

# Calculate HLNM percentiles for PiP areas
Write-Host "  Calculating HLNM percentiles..."
$hlnmFields = @("growth", "energy", "crime", "opportunity", "health")
foreach ($field in $hlnmFields) {
    $allValues = @($allHlnm.Values | ForEach-Object { $_.$field } | Sort-Object)
    foreach ($code in $msoaData.Keys) {
        if ($msoaData[$code]["hlnm_$field"]) {
            $val = $msoaData[$code]["hlnm_$field"]
            $rank = ($allValues | Where-Object { $_ -lt $val }).Count
            $percentile = [Math]::Round(($rank / $allValues.Count) * 100, 1)
            # For crime, lower is worse, so invert
            if ($field -eq "crime") {
                $percentile = 100 - $percentile
            }
            $msoaData[$code]["hlnm_${field}_percentile"] = $percentile
        }
    }
}

# ============================================
# 3. Read Community Needs Index files
# ============================================
Write-Host "3. Reading Community Needs Index files..."
$cniFiles = @{
    "cni_overall_rank" = "MSOA_Community Needs Index 2023_ Community Needs rank_2023-01-01.xlsx"
    "cni_active_engaged_rank" = "MSOA_Community Needs Index 2023_ Active and Engaged Community rank_2023-01-01.xlsx"
    "cni_civic_assets_rank" = "MSOA_Community Needs Index 2023_ Civic Assets rank_2023-01-01.xlsx"
    "cni_connectedness_rank" = "MSOA_Community Needs Index 2023_ Connectedness rank_2023-01-01.xlsx"
}

$allCni = @{}
foreach ($cniKey in $cniFiles.Keys) {
    $cniPath = Join-Path $basePath $cniFiles[$cniKey]
    Write-Host "  Reading $cniKey..."
    $cniWorkbook = $excel.Workbooks.Open($cniPath)
    $dataSheet = $cniWorkbook.Sheets.Item("Data")
    $usedRange = $dataSheet.UsedRange

    for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
        $areaCode = $dataSheet.Cells.Item($row, 1).Text
        if ($areaCode -match "^E02") {
            $value = [int]$dataSheet.Cells.Item($row, 3).Value2
            if (-not $allCni.ContainsKey($areaCode)) { $allCni[$areaCode] = @{} }
            $allCni[$areaCode][$cniKey] = $value
            if ($msoaData.ContainsKey($areaCode)) {
                $msoaData[$areaCode][$cniKey] = $value
            }
        }
    }
    $cniWorkbook.Close($false)
}
Write-Host "  Read CNI data for $($allCni.Count) MSOAs"

# Calculate CNI percentiles (lower rank = higher need)
Write-Host "  Calculating CNI percentiles..."
$totalMsoas = $allCni.Count
foreach ($code in $msoaData.Keys) {
    foreach ($cniKey in $cniFiles.Keys) {
        if ($msoaData[$code][$cniKey]) {
            $rank = $msoaData[$code][$cniKey]
            # Convert rank to percentile (rank 1 = 100th percentile of need)
            $percentile = [Math]::Round((1 - ($rank / $totalMsoas)) * 100, 1)
            $msoaData[$code]["${cniKey}_percentile"] = $percentile
        }
    }
}

# ============================================
# 4. Calculate benchmarks
# ============================================
Write-Host "4. Calculating benchmarks..."
$benchmarks = @{
    total_msoas = $allHlnm.Count
    hlnm_growth_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.growth } | Measure-Object -Average).Average, 2)
    hlnm_energy_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.energy } | Measure-Object -Average).Average, 2)
    hlnm_crime_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.crime } | Measure-Object -Average).Average, 2)
    hlnm_opportunity_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.opportunity } | Measure-Object -Average).Average, 2)
    hlnm_health_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.health } | Measure-Object -Average).Average, 2)
    cni_median_rank = [Math]::Round($totalMsoas / 2)
}

# Calculate LA-level benchmarks
Write-Host "  Calculating LA benchmarks..."
$laStats = @{}
foreach ($code in $allHlnm.Keys) {
    # Group by LA (first 3 digits of MSOA code don't help, need to use CNI parent data)
    # For now, calculate from PiP areas' LAs
}

# ============================================
# 5. Export to JSON
# ============================================
Write-Host "5. Exporting to JSON..."

$output = @{
    metadata = @{
        generated = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        total_pip_areas = $msoaData.Count
        total_msoas_england = $allHlnm.Count
        data_sources = @{
            pride_in_place = "OCSI - Pride in Place Programme neighbourhoods"
            census_2021 = "ONS Census 2021 via NOMIS"
            imd_2025 = "MHCLG English Indices of Deprivation 2025"
            hlnm = "OCSI - Hyper-Local Need Measure 2025"
            cni = "OCSI/Local Trust - Community Needs Index 2023"
        }
        citations = @{
            hlnm = @{
                source = "Oxford Consultants for Social Inclusion (OCSI)"
                title = "Hyper-Local Need Measure 2025"
                url = "https://ocsi.uk"
                note = "Higher values indicate greater need (except Crime where lower = worse)"
            }
            cni = @{
                source = "Oxford Consultants for Social Inclusion (OCSI) for Local Trust"
                title = "Community Needs Index 2023"
                url = "https://localtrust.org.uk/policy/left-behind-neighbourhoods/"
                note = "Lower rank indicates higher need"
            }
            census = @{
                source = "Office for National Statistics"
                title = "Census 2021"
                url = "https://www.nomisweb.co.uk"
            }
            imd = @{
                source = "Ministry of Housing, Communities and Local Government"
                title = "English Indices of Deprivation 2025"
                url = "https://www.gov.uk/government/statistics/english-indices-of-deprivation-2025"
            }
        }
    }
    benchmarks = $benchmarks
    areas = @($msoaData.Values)
}

$jsonOutput = $output | ConvertTo-Json -Depth 10 -Compress:$false
$outputPath = Join-Path $basePath "data_comprehensive.json"
[System.IO.File]::WriteAllText($outputPath, $jsonOutput, [System.Text.Encoding]::UTF8)

Write-Host "  Exported to: $outputPath"
Write-Host "  File size: $([Math]::Round((Get-Item $outputPath).Length / 1KB, 2)) KB"

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host ""
Write-Host "=== MSOA Data extraction complete! ==="
