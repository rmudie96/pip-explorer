$ErrorActionPreference = "Stop"
$basePath = "C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer"

Write-Host "=== Pride in Place Complete Data Extraction ==="
Write-Host "Started: $(Get-Date -Format 'HH:mm:ss')"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# ============================================
# 1. Read existing data.json to get PiP MSOAs
# ============================================
Write-Host ""
Write-Host "1. Reading existing data.json..."
$existingData = Get-Content (Join-Path $basePath "data.json") -Raw | ConvertFrom-Json
$msoaData = @{}
$pipMsoaNames = @{}
foreach ($area in $existingData.areas) {
    $msoaData[$area.msoa_code] = @{}
    foreach ($prop in $area.PSObject.Properties) {
        $msoaData[$area.msoa_code][$prop.Name] = $prop.Value
    }
    $msoaData[$area.msoa_code]["lsoas"] = @()
    $pipMsoaNames[$area.neighbourhood_name] = $area.msoa_code
}
Write-Host "  Found $($msoaData.Count) Pride in Place areas"

# ============================================
# 2. Read HLNM_MSOA.xlsx
# ============================================
Write-Host ""
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
            $msoaData[$areaCode]["hlnm_growth"] = [Math]::Round($allHlnm[$areaCode].growth, 2)
            $msoaData[$areaCode]["hlnm_energy"] = [Math]::Round($allHlnm[$areaCode].energy, 2)
            $msoaData[$areaCode]["hlnm_crime"] = [Math]::Round($allHlnm[$areaCode].crime, 0)
            $msoaData[$areaCode]["hlnm_opportunity"] = [Math]::Round($allHlnm[$areaCode].opportunity, 2)
            $msoaData[$areaCode]["hlnm_health"] = [Math]::Round($allHlnm[$areaCode].health, 2)
        }
    }
}
Write-Host "  Read $($allHlnm.Count) MSOA HLNM records"
$hlnmWorkbook.Close($false)

# Calculate HLNM percentiles
Write-Host "  Calculating HLNM percentiles..."
$hlnmFields = @("growth", "energy", "crime", "opportunity", "health")
foreach ($field in $hlnmFields) {
    $allValues = @($allHlnm.Values | ForEach-Object { $_.$field } | Sort-Object)
    foreach ($code in $msoaData.Keys) {
        if ($null -ne $msoaData[$code]["hlnm_$field"]) {
            $val = $allHlnm[$code].$field
            $rank = ($allValues | Where-Object { $_ -lt $val }).Count
            $percentile = [Math]::Round(($rank / $allValues.Count) * 100, 1)
            # For crime, lower raw value = worse, so DON'T invert (higher percentile = more deprived already)
            # For all others, higher raw value = worse (higher percentile = more deprived)
            if ($field -eq "crime") {
                # Lower crime score = worse, so invert so higher percentile = more deprived
                $percentile = [Math]::Round(100 - $percentile, 1)
            }
            $msoaData[$code]["hlnm_${field}_percentile"] = $percentile
        }
    }
}

# ============================================
# 3. Read Community Needs Index files
# ============================================
Write-Host ""
Write-Host "3. Reading Community Needs Index files..."
$cniFiles = @{
    "cni_overall_rank" = "MSOA_Community Needs Index 2023_ Community Needs rank_2023-01-01.xlsx"
    "cni_active_engaged_rank" = "MSOA_Community Needs Index 2023_ Active and Engaged Community rank_2023-01-01.xlsx"
    "cni_civic_assets_rank" = "MSOA_Community Needs Index 2023_ Civic Assets rank_2023-01-01.xlsx"
    "cni_connectedness_rank" = "MSOA_Community Needs Index 2023_ Connectedness rank_2023-01-01.xlsx"
}

# CNI ranks are out of ~33,755 (all LSOAs)
$cniTotalLsoas = 33755

foreach ($cniKey in $cniFiles.Keys) {
    $cniPath = Join-Path $basePath $cniFiles[$cniKey]
    Write-Host "  Reading $cniKey..."
    $cniWorkbook = $excel.Workbooks.Open($cniPath)
    $dataSheet = $cniWorkbook.Sheets.Item("Data")
    $usedRange = $dataSheet.UsedRange

    for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
        $areaCode = $dataSheet.Cells.Item($row, 1).Text
        if ($areaCode -match "^E02" -and $msoaData.ContainsKey($areaCode)) {
            $rank = [int]$dataSheet.Cells.Item($row, 3).Value2
            $msoaData[$areaCode][$cniKey] = $rank
            # Convert rank to percentile (rank 1 = highest need = 100th percentile)
            $percentile = [Math]::Round((1 - ($rank / $cniTotalLsoas)) * 100, 1)
            $msoaData[$areaCode]["${cniKey}_percentile"] = $percentile
        }
    }
    $cniWorkbook.Close($false)
}
Write-Host "  Done reading CNI files"

# ============================================
# 4. Read LSOA-level data (targeted for PiP areas only)
# ============================================
Write-Host ""
Write-Host "4. Reading LSOA-level data..."
$lsoaPath = Join-Path $basePath "Hyper-Local-Need-Measure-2025.xlsx"
$lsoaWorkbook = $excel.Workbooks.Open($lsoaPath)
$lsoaSheet = $lsoaWorkbook.Sheets.Item(1)
$usedRange = $lsoaSheet.UsedRange

$lsoaCount = 0
$matchedCount = 0

for ($row = 6; $row -le $usedRange.Rows.Count; $row++) {
    $lsoaCode = $lsoaSheet.Cells.Item($row, 1).Text
    if ($lsoaCode -match "^E01") {
        $lsoaCount++
        $lsoaName = $lsoaSheet.Cells.Item($row, 2).Text
        $laName = $lsoaSheet.Cells.Item($row, 4).Text
        $hlnmScore = $lsoaSheet.Cells.Item($row, 5).Value2
        $typology = $lsoaSheet.Cells.Item($row, 10).Text

        # Try to match to a PiP MSOA by checking if LSOA name contains MSOA neighbourhood name
        foreach ($msoaName in $pipMsoaNames.Keys) {
            # Match pattern: LSOA name contains the MSOA name followed by a space and 3 digits
            if ($lsoaName -match [regex]::Escape($msoaName) -and $lsoaName -match "\s\d{3}[A-Z]$") {
                $msoaCode = $pipMsoaNames[$msoaName]
                # Also verify LA matches
                if ($msoaData[$msoaCode].local_authority -eq $laName) {
                    $lsoaRecord = @{
                        lsoa_code = $lsoaCode
                        lsoa_name = $lsoaName
                        hlnm_score = [Math]::Round($hlnmScore, 2)
                        typology = $typology
                    }
                    $msoaData[$msoaCode]["lsoas"] += $lsoaRecord
                    $matchedCount++
                    break
                }
            }
        }
    }
    if ($row % 10000 -eq 0) {
        Write-Host "    Processed $row rows, matched $matchedCount LSOAs..."
    }
}
Write-Host "  Read $lsoaCount total LSOAs, matched $matchedCount to PiP areas"
$lsoaWorkbook.Close($false)

# Calculate LSOA summary stats for each MSOA
Write-Host "  Calculating LSOA summaries..."
foreach ($code in $msoaData.Keys) {
    $lsoas = $msoaData[$code]["lsoas"]
    $msoaData[$code]["lsoa_count"] = $lsoas.Count
    $msoaData[$code]["lsoa_mission_critical"] = @($lsoas | Where-Object { $_.typology -eq "Mission Critical" }).Count
    $msoaData[$code]["lsoa_mission_priority"] = @($lsoas | Where-Object { $_.typology -eq "Mission Priority" }).Count
    $msoaData[$code]["lsoa_mission_support"] = @($lsoas | Where-Object { $_.typology -eq "Mission Support" }).Count
}

# ============================================
# 5. Calculate benchmarks
# ============================================
Write-Host ""
Write-Host "5. Calculating benchmarks..."
$benchmarks = @{
    total_msoas_england = $allHlnm.Count
    total_lsoas_england = $cniTotalLsoas
    hlnm_growth_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.growth } | Measure-Object -Average).Average, 2)
    hlnm_energy_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.energy } | Measure-Object -Average).Average, 2)
    hlnm_crime_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.crime } | Measure-Object -Average).Average, 0)
    hlnm_opportunity_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.opportunity } | Measure-Object -Average).Average, 2)
    hlnm_health_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.health } | Measure-Object -Average).Average, 2)
}

# ============================================
# 6. Build LA summary data
# ============================================
Write-Host ""
Write-Host "6. Building LA summary..."
$laData = @{}
foreach ($code in $msoaData.Keys) {
    $la = $msoaData[$code].local_authority
    if (-not $laData.ContainsKey($la)) {
        $laData[$la] = @{
            name = $la
            msoa_codes = @()
            pip_area_count = 0
        }
    }
    $laData[$la].msoa_codes += $code
    $laData[$la].pip_area_count++
}

# ============================================
# 7. Export to JSON
# ============================================
Write-Host ""
Write-Host "7. Exporting to JSON..."

$output = @{
    metadata = @{
        generated = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        total_pip_areas = $msoaData.Count
        total_msoas_england = $allHlnm.Count
        total_lsoas_england = $cniTotalLsoas
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
                note = "Higher values indicate greater need (except Crime where lower values indicate greater need)"
            }
            cni = @{
                source = "Oxford Consultants for Social Inclusion (OCSI) for Local Trust"
                title = "Community Needs Index 2023"
                url = "https://localtrust.org.uk/policy/left-behind-neighbourhoods/"
                note = "Lower rank indicates higher need (rank 1 = highest need)"
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
    local_authorities = @($laData.Values)
}

$jsonOutput = $output | ConvertTo-Json -Depth 10 -Compress:$false
$outputPath = Join-Path $basePath "data_comprehensive.json"
[System.IO.File]::WriteAllText($outputPath, $jsonOutput, [System.Text.Encoding]::UTF8)

Write-Host "  Exported to: $outputPath"
Write-Host "  File size: $([Math]::Round((Get-Item $outputPath).Length / 1KB, 2)) KB"

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host ""
Write-Host "=== Complete data extraction finished at $(Get-Date -Format 'HH:mm:ss')! ==="
