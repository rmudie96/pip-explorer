$ErrorActionPreference = "Stop"
$basePath = "C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer"

Write-Host "=== Building Pride in Place Data ==="
Write-Host "Started: $(Get-Date -Format 'HH:mm:ss')"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # 1. Load existing PiP areas
    Write-Host "`n1. Loading existing PiP areas..."
    $existing = Get-Content "$basePath\data.json" -Raw | ConvertFrom-Json
    $pipMsoas = @{}
    foreach ($a in $existing.areas) {
        $pipMsoas[$a.msoa_code] = @{
            msoa_code = $a.msoa_code
            neighbourhood_name = $a.neighbourhood_name
            local_authority = $a.local_authority
        }
        foreach ($p in $a.PSObject.Properties) {
            $pipMsoas[$a.msoa_code][$p.Name] = $p.Value
        }
    }
    Write-Host "  Loaded $($pipMsoas.Count) PiP areas"

    # 2. Load lookup and filter to PiP MSOAs
    Write-Host "`n2. Loading lookup table..."
    $wb = $excel.Workbooks.Open("$basePath\lsoa_msoa_la_region_lookup.xlsx")
    $sheet = $wb.Sheets.Item(1)
    $rows = $sheet.UsedRange.Rows.Count

    $lsoaToMsoa = @{}
    $msoaRegion = @{}
    for ($r = 2; $r -le $rows; $r++) {
        $lsoa = $sheet.Cells.Item($r, 1).Text
        $msoa = $sheet.Cells.Item($r, 4).Text
        $region = $sheet.Cells.Item($r, 10).Text
        $lsoaToMsoa[$lsoa] = $msoa
        if ($pipMsoas.ContainsKey($msoa)) {
            $msoaRegion[$msoa] = $region
        }
    }
    $wb.Close($false)
    Write-Host "  Loaded $($lsoaToMsoa.Count) LSOA mappings"

    # Add region to PiP areas
    foreach ($code in $msoaRegion.Keys) {
        $pipMsoas[$code]["region"] = $msoaRegion[$code]
    }

    # 3. Load HLNM MSOA data
    Write-Host "`n3. Loading HLNM MSOA data..."
    $wb = $excel.Workbooks.Open("$basePath\HLNM_MSOA.xlsx")
    $sheet = $wb.Sheets.Item("master")
    $rows = $sheet.UsedRange.Rows.Count

    $allHlnm = @{}
    for ($r = 2; $r -le $rows; $r++) {
        $code = $sheet.Cells.Item($r, 1).Text
        if ($code -match "^E02") {
            $allHlnm[$code] = @{
                growth = [double]$sheet.Cells.Item($r, 5).Value2
                energy = [double]$sheet.Cells.Item($r, 6).Value2
                crime = [double]$sheet.Cells.Item($r, 7).Value2
                opportunity = [double]$sheet.Cells.Item($r, 8).Value2
                health = [double]$sheet.Cells.Item($r, 9).Value2
            }
        }
    }
    $wb.Close($false)
    Write-Host "  Loaded $($allHlnm.Count) MSOA HLNM records"

    # Add HLNM to PiP areas and calculate percentiles
    foreach ($code in $pipMsoas.Keys) {
        if ($allHlnm.ContainsKey($code)) {
            $pipMsoas[$code]["hlnm_growth"] = [Math]::Round($allHlnm[$code].growth, 2)
            $pipMsoas[$code]["hlnm_energy"] = [Math]::Round($allHlnm[$code].energy, 2)
            $pipMsoas[$code]["hlnm_crime"] = [Math]::Round($allHlnm[$code].crime, 0)
            $pipMsoas[$code]["hlnm_opportunity"] = [Math]::Round($allHlnm[$code].opportunity, 2)
            $pipMsoas[$code]["hlnm_health"] = [Math]::Round($allHlnm[$code].health, 2)
        }
    }

    # Calculate percentiles
    $fields = @("growth", "energy", "crime", "opportunity", "health")
    foreach ($f in $fields) {
        $sorted = @($allHlnm.Values | ForEach-Object { $_.$f } | Sort-Object)
        foreach ($code in $pipMsoas.Keys) {
            if ($pipMsoas[$code]["hlnm_$f"]) {
                $val = $allHlnm[$code].$f
                $below = ($sorted | Where-Object { $_ -lt $val }).Count
                $pct = [Math]::Round(($below / $sorted.Count) * 100, 1)
                if ($f -eq "crime") { $pct = 100 - $pct }
                $pipMsoas[$code]["hlnm_${f}_percentile"] = $pct
            }
        }
    }

    # 4. Load CNI data
    Write-Host "`n4. Loading CNI data..."
    $cniFiles = @{
        "cni_overall_rank" = "MSOA_Community Needs Index 2023_ Community Needs rank_2023-01-01.xlsx"
        "cni_active_engaged_rank" = "MSOA_Community Needs Index 2023_ Active and Engaged Community rank_2023-01-01.xlsx"
        "cni_civic_assets_rank" = "MSOA_Community Needs Index 2023_ Civic Assets rank_2023-01-01.xlsx"
        "cni_connectedness_rank" = "MSOA_Community Needs Index 2023_ Connectedness rank_2023-01-01.xlsx"
    }
    $cniTotal = 33755

    foreach ($key in $cniFiles.Keys) {
        $wb = $excel.Workbooks.Open("$basePath\$($cniFiles[$key])")
        $sheet = $wb.Sheets.Item("Data")
        $rows = $sheet.UsedRange.Rows.Count
        for ($r = 2; $r -le $rows; $r++) {
            $code = $sheet.Cells.Item($r, 1).Text
            if ($pipMsoas.ContainsKey($code)) {
                $rank = [int]$sheet.Cells.Item($r, 3).Value2
                $pipMsoas[$code][$key] = $rank
                $pipMsoas[$code]["${key}_percentile"] = [Math]::Round((1 - $rank / $cniTotal) * 100, 1)
            }
        }
        $wb.Close($false)
    }
    Write-Host "  Loaded CNI data"

    # 5. Load LSOA HLNM data and match to PiP MSOAs
    Write-Host "`n5. Loading LSOA HLNM data..."
    $wb = $excel.Workbooks.Open("$basePath\Hyper-Local-Need-Measure-2025.xlsx")
    $sheet = $wb.Sheets.Item(1)
    $rows = $sheet.UsedRange.Rows.Count

    # Initialize LSOA arrays
    foreach ($code in $pipMsoas.Keys) {
        $pipMsoas[$code]["lsoas"] = @()
        $pipMsoas[$code]["lsoa_mission_critical"] = 0
        $pipMsoas[$code]["lsoa_mission_priority"] = 0
        $pipMsoas[$code]["lsoa_mission_support"] = 0
    }

    $matched = 0
    for ($r = 6; $r -le $rows; $r++) {
        $lsoa = $sheet.Cells.Item($r, 1).Text
        if ($lsoa -match "^E01" -and $lsoaToMsoa.ContainsKey($lsoa)) {
            $msoa = $lsoaToMsoa[$lsoa]
            if ($pipMsoas.ContainsKey($msoa)) {
                $name = $sheet.Cells.Item($r, 2).Text
                $score = [Math]::Round([double]$sheet.Cells.Item($r, 5).Value2, 2)
                $typology = $sheet.Cells.Item($r, 10).Text

                $pipMsoas[$msoa]["lsoas"] += @{
                    lsoa_code = $lsoa
                    lsoa_name = $name
                    hlnm_score = $score
                    typology = $typology
                }

                switch ($typology) {
                    "Mission Critical" { $pipMsoas[$msoa]["lsoa_mission_critical"]++ }
                    "Mission Priority" { $pipMsoas[$msoa]["lsoa_mission_priority"]++ }
                    "Mission Support" { $pipMsoas[$msoa]["lsoa_mission_support"]++ }
                }
                $matched++
            }
        }
        if ($r % 10000 -eq 0) { Write-Host "    Row $r, matched $matched" }
    }
    $wb.Close($false)
    Write-Host "  Matched $matched LSOAs to PiP areas"

    # 6. Calculate benchmarks
    Write-Host "`n6. Calculating benchmarks..."
    $benchmarks = @{
        total_msoas = $allHlnm.Count
        total_lsoas = $cniTotal
        hlnm_growth_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.growth } | Measure-Object -Average).Average, 2)
        hlnm_energy_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.energy } | Measure-Object -Average).Average, 2)
        hlnm_crime_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.crime } | Measure-Object -Average).Average, 0)
        hlnm_opportunity_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.opportunity } | Measure-Object -Average).Average, 2)
        hlnm_health_avg = [Math]::Round(($allHlnm.Values | ForEach-Object { $_.health } | Measure-Object -Average).Average, 2)
    }

    # 7. Build LA summary
    Write-Host "`n7. Building LA summary..."
    $laData = @{}
    foreach ($code in $pipMsoas.Keys) {
        $la = $pipMsoas[$code].local_authority
        if (-not $laData.ContainsKey($la)) {
            $laData[$la] = @{ name = $la; msoa_codes = @(); pip_count = 0 }
        }
        $laData[$la].msoa_codes += $code
        $laData[$la].pip_count++
    }

    # 8. Export
    Write-Host "`n8. Exporting JSON..."
    $output = @{
        metadata = @{
            generated = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            total_pip_areas = $pipMsoas.Count
        }
        benchmarks = $benchmarks
        areas = @($pipMsoas.Values)
        local_authorities = @($laData.Values)
    }

    $json = $output | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText("$basePath\data_final.json", $json, [System.Text.Encoding]::UTF8)
    Write-Host "  Saved to data_final.json"

} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "`n=== Done at $(Get-Date -Format 'HH:mm:ss') ==="
