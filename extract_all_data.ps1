$ErrorActionPreference = "Stop"
$basePath = "C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer"

Write-Host "=== Pride in Place Data Extraction ==="
Write-Host "Starting Excel COM automation..."

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Initialize data structures
$msoaData = @{}
$lsoaData = @{}
$laData = @{}

# ============================================
# 1. Read existing data.json to get PiP MSOAs
# ============================================
Write-Host ""
Write-Host "1. Reading existing data.json..."
$existingData = Get-Content (Join-Path $basePath "data.json") -Raw | ConvertFrom-Json
$pipMsoas = @{}
foreach ($area in $existingData.areas) {
    $pipMsoas[$area.msoa_code] = $area
    $msoaData[$area.msoa_code] = @{
        msoa_code = $area.msoa_code
        neighbourhood_name = $area.neighbourhood_name
        local_authority = $area.local_authority
    }
    # Copy existing data
    foreach ($prop in $area.PSObject.Properties) {
        $msoaData[$area.msoa_code][$prop.Name] = $prop.Value
    }
}
Write-Host "  Found $($pipMsoas.Count) Pride in Place areas"

# ============================================
# 2. Read HLNM_MSOA.xlsx
# ============================================
Write-Host ""
Write-Host "2. Reading HLNM_MSOA.xlsx..."
$hlnmPath = Join-Path $basePath "HLNM_MSOA.xlsx"
$hlnmWorkbook = $excel.Workbooks.Open($hlnmPath)
$masterSheet = $hlnmWorkbook.Sheets.Item("master")
$usedRange = $masterSheet.UsedRange

# Read all HLNM data (for benchmarking)
$allHlnmData = @{}
for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
    $areaCode = $masterSheet.Cells.Item($row, 1).Text
    if ($areaCode -match "^E02") {
        $allHlnmData[$areaCode] = @{
            growth = [double]$masterSheet.Cells.Item($row, 5).Value2
            energy = [double]$masterSheet.Cells.Item($row, 6).Value2
            crime = [double]$masterSheet.Cells.Item($row, 7).Value2
            opportunity = [double]$masterSheet.Cells.Item($row, 8).Value2
            health = [double]$masterSheet.Cells.Item($row, 9).Value2
        }

        # If it's a PiP MSOA, add to our data
        if ($msoaData.ContainsKey($areaCode)) {
            $msoaData[$areaCode]["hlnm_growth"] = $allHlnmData[$areaCode].growth
            $msoaData[$areaCode]["hlnm_energy"] = $allHlnmData[$areaCode].energy
            $msoaData[$areaCode]["hlnm_crime"] = $allHlnmData[$areaCode].crime
            $msoaData[$areaCode]["hlnm_opportunity"] = $allHlnmData[$areaCode].opportunity
            $msoaData[$areaCode]["hlnm_health"] = $allHlnmData[$areaCode].health
        }
    }
}
Write-Host "  Read $($allHlnmData.Count) MSOA HLNM records"
$hlnmWorkbook.Close($false)

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

$allCniData = @{}
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

            if (-not $allCniData.ContainsKey($areaCode)) {
                $allCniData[$areaCode] = @{}
            }
            $allCniData[$areaCode][$cniKey] = $value

            # If it's a PiP MSOA, add to our data
            if ($msoaData.ContainsKey($areaCode)) {
                $msoaData[$areaCode][$cniKey] = $value
            }
        }
    }
    $cniWorkbook.Close($false)
}
Write-Host "  Read CNI data for $($allCniData.Count) MSOAs"

# ============================================
# 4. Read LSOA-level HLNM data
# ============================================
Write-Host ""
Write-Host "4. Reading LSOA-level HLNM data..."
$lsoaPath = Join-Path $basePath "Hyper-Local-Need-Measure-2025.xlsx"
$lsoaWorkbook = $excel.Workbooks.Open($lsoaPath)
$lsoaSheet = $lsoaWorkbook.Sheets.Item(1)
$usedRange = $lsoaSheet.UsedRange

# We need LSOA to MSOA mapping - read from data or build it
# For now, store all LSOA data and we'll link via name matching
$allLsoaData = @()
$msoaLsoaSummary = @{}

for ($row = 6; $row -le $usedRange.Rows.Count; $row++) {
    $lsoaCode = $lsoaSheet.Cells.Item($row, 1).Text
    if ($lsoaCode -match "^E01") {
        $lsoaName = $lsoaSheet.Cells.Item($row, 2).Text
        $laCode = $lsoaSheet.Cells.Item($row, 3).Text
        $laName = $lsoaSheet.Cells.Item($row, 4).Text
        $hlnmScore = $lsoaSheet.Cells.Item($row, 5).Value2
        $typology = $lsoaSheet.Cells.Item($row, 10).Text

        # Extract MSOA name from LSOA name (format: "Area Name - MSOA Name XXXZ")
        $msoaName = ""
        if ($lsoaName -match "^(.+)\s+-\s+(.+)\s+\d{3}[A-Z]$") {
            $msoaName = $matches[2]
        }

        $lsoaRecord = @{
            lsoa_code = $lsoaCode
            lsoa_name = $lsoaName
            msoa_name = $msoaName
            la_code = $laCode
            la_name = $laName
            hlnm_score = $hlnmScore
            typology = $typology
        }
        $allLsoaData += $lsoaRecord

        # Track LA data
        if (-not $laData.ContainsKey($laCode)) {
            $laData[$laCode] = @{
                la_code = $laCode
                la_name = $laName
                msoas = @()
                lsoa_count = 0
                mission_critical = 0
                mission_priority = 0
                mission_support = 0
            }
        }
        $laData[$laCode].lsoa_count++
        if ($typology -eq "Mission Critical") { $laData[$laCode].mission_critical++ }
        elseif ($typology -eq "Mission Priority") { $laData[$laCode].mission_priority++ }
        elseif ($typology -eq "Mission Support") { $laData[$laCode].mission_support++ }
    }

    if ($row % 5000 -eq 0) {
        Write-Host "    Processed $row rows..."
    }
}
Write-Host "  Read $($allLsoaData.Count) LSOA records"
$lsoaWorkbook.Close($false)

# ============================================
# 5. Link LSOAs to MSOAs for PiP areas
# ============================================
Write-Host ""
Write-Host "5. Linking LSOAs to MSOAs..."

# Build MSOA name to code lookup from our PiP data
$msoaNameLookup = @{}
foreach ($code in $msoaData.Keys) {
    $name = $msoaData[$code].neighbourhood_name
    if ($name) {
        $msoaNameLookup[$name] = $code
    }
}

# Group LSOAs by matching to PiP MSOAs
foreach ($code in $msoaData.Keys) {
    $msoaData[$code]["lsoas"] = @()
    $msoaData[$code]["lsoa_mission_critical"] = 0
    $msoaData[$code]["lsoa_mission_priority"] = 0
    $msoaData[$code]["lsoa_mission_support"] = 0
    $msoaData[$code]["lsoa_other"] = 0
}

# Match LSOAs to MSOAs by name pattern
foreach ($lsoa in $allLsoaData) {
    foreach ($code in $msoaData.Keys) {
        $msoaName = $msoaData[$code].neighbourhood_name
        if ($msoaName -and $lsoa.lsoa_name -like "*$msoaName*") {
            $msoaData[$code]["lsoas"] += $lsoa
            switch ($lsoa.typology) {
                "Mission Critical" { $msoaData[$code]["lsoa_mission_critical"]++ }
                "Mission Priority" { $msoaData[$code]["lsoa_mission_priority"]++ }
                "Mission Support" { $msoaData[$code]["lsoa_mission_support"]++ }
                default { $msoaData[$code]["lsoa_other"]++ }
            }
            break
        }
    }
}

# ============================================
# 6. Calculate benchmarks (national averages)
# ============================================
Write-Host ""
Write-Host "6. Calculating national benchmarks..."

$benchmarks = @{
    hlnm_growth_avg = ($allHlnmData.Values | ForEach-Object { $_.growth } | Measure-Object -Average).Average
    hlnm_energy_avg = ($allHlnmData.Values | ForEach-Object { $_.energy } | Measure-Object -Average).Average
    hlnm_crime_avg = ($allHlnmData.Values | ForEach-Object { $_.crime } | Measure-Object -Average).Average
    hlnm_opportunity_avg = ($allHlnmData.Values | ForEach-Object { $_.opportunity } | Measure-Object -Average).Average
    hlnm_health_avg = ($allHlnmData.Values | ForEach-Object { $_.health } | Measure-Object -Average).Average
    cni_overall_rank_median = ($allCniData.Values | ForEach-Object { $_.cni_overall_rank } | Sort-Object | Select-Object -Index ([Math]::Floor($allCniData.Count / 2)))
    total_msoas = $allHlnmData.Count
}
Write-Host "  National HLNM Growth avg: $([Math]::Round($benchmarks.hlnm_growth_avg, 2))"
Write-Host "  National HLNM Crime avg: $([Math]::Round($benchmarks.hlnm_crime_avg, 2))"

# ============================================
# 7. Export to JSON
# ============================================
Write-Host ""
Write-Host "7. Exporting to JSON..."

$output = @{
    metadata = @{
        generated = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        total_pip_areas = $msoaData.Count
        total_msoas_england = $allHlnmData.Count
        data_sources = @{
            pride_in_place = "OCSI - Pride in Place Programme neighbourhoods"
            census_2021 = "ONS Census 2021 via NOMIS"
            imd_2025 = "MHCLG English Indices of Deprivation 2025"
            hlnm = "OCSI - Hyper-Local Need Measure 2025"
            cni = "OCSI - Community Needs Index 2023"
        }
        citations = @{
            hlnm = "Oxford Consultants for Social Inclusion (OCSI). Hyper-Local Need Measure 2025. https://ocsi.uk"
            cni = "Oxford Consultants for Social Inclusion (OCSI) for Local Trust. Community Needs Index 2023. https://localtrust.org.uk"
            census = "Office for National Statistics. Census 2021. https://www.nomisweb.co.uk"
            imd = "Ministry of Housing, Communities & Local Government. English Indices of Deprivation 2025. https://www.gov.uk"
        }
    }
    benchmarks = $benchmarks
    areas = @($msoaData.Values)
    local_authorities = @($laData.Values | Where-Object { $_.msoas.Count -gt 0 -or $_.mission_critical -gt 0 })
}

$jsonOutput = $output | ConvertTo-Json -Depth 10
$outputPath = Join-Path $basePath "data_comprehensive.json"
$jsonOutput | Out-File -FilePath $outputPath -Encoding UTF8

Write-Host "  Exported to: $outputPath"
Write-Host "  File size: $([Math]::Round((Get-Item $outputPath).Length / 1KB, 2)) KB"

# ============================================
# Cleanup
# ============================================
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host ""
Write-Host "=== Data extraction complete! ==="
