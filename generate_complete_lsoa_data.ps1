# Generate complete LSOA datasets for all 191 LSOAs across 40 Pride in Place MSOAs
# Outputs: lsoa_hlnm_data_complete.js, lsoa_economic_data_complete.js, lsoa_map_data_complete.js

Write-Host "`n=== Building Complete Pride in Place LSOA Dataset ===" -ForegroundColor Cyan
Write-Host "This will process all 191 LSOAs across 40 MSOAs`n"

# 1. Get Pride in Place MSOA list
$htmlContent = Get-Content "index.html" -Raw
$dataStart = $htmlContent.IndexOf('const DATA = ') + 13
$dataEnd = $htmlContent.IndexOf('};', $dataStart) + 1
$jsonStr = $htmlContent.Substring($dataStart, $dataEnd - $dataStart)
$data = $jsonStr | ConvertFrom-Json
$pipMSOAs = $data.areas | Select-Object -ExpandProperty msoa_code

# 2. Get all LSOAs for PiP MSOAs
$lookup = Import-Csv "lsoa_msoa_la_region_lookup_csv.csv"
$pipLSOAs = $lookup | Where-Object { $pipMSOAs -contains $_.MSOA21CD }

Write-Host "Step 1: Loading HLNM data..." -ForegroundColor Yellow
$hlnmData = Import-Csv "Hyper-Local-Need-Measure-2025-_csv.csv"
$hlnmByCode = @{}
foreach ($row in $hlnmData) {
    $hlnmByCode[$row.'Area Code'] = $row
}
Write-Host "  Loaded HLNM data for $($hlnmByCode.Count) LSOAs"

Write-Host "`nStep 2: Loading economic data..." -ForegroundColor Yellow
$econData = Import-Csv "Hyper-local Need Index_econ_underlying.csv"
$econByCode = @{}
foreach ($row in $econData) {
    $econByCode[$row.'Area Code'] = $row
}
Write-Host "  Loaded economic data for $($econByCode.Count) LSOAs"

Write-Host "`nStep 3: Processing GeoJSON for coordinates..." -ForegroundColor Yellow
Write-Host "  Loading GeoJSON file (this may take a moment)..."
$geojsonContent = Get-Content "Lower_layer_Super_Output_Areas_December_2021_Boundaries_EW_BSC_V4_-4299016806856585929.geojson" -Raw | ConvertFrom-Json
Write-Host "  Loaded $($geojsonContent.features.Count) LSOA boundaries"

# Build coordinate lookup
$coordsByCode = @{}
foreach ($feature in $geojsonContent.features) {
    $code = $feature.properties.LSOA21CD
    # Get centroid from coordinates (simplified - uses first point)
    if ($feature.geometry.coordinates) {
        try {
            $coords = $feature.geometry.coordinates
            # GeoJSON is [lng, lat] format and can be nested
            if ($coords[0][0] -is [Array]) {
                $lng = [math]::Round(($coords[0][0][0] -as [double]), 5)
                $lat = [math]::Round(($coords[0][0][1] -as [double]), 5)
            } else {
                $lng = [math]::Round(($coords[0][0] -as [double]), 5)
                $lat = [math]::Round(($coords[0][1] -as [double]), 5)
            }
            $coordsByCode[$code] = @{ lat = $lat; lng = $lng }
        } catch {
            Write-Host "  Warning: Could not parse coords for $code" -ForegroundColor DarkYellow
        }
    }
}
Write-Host "  Extracted coordinates for $($coordsByCode.Count) LSOAs"

Write-Host "`nStep 4: Building output datasets..." -ForegroundColor Yellow

# Build HLNM JS file
$hlnmJS = "const LSOA_HLNM_DATA = {"
$hlnmCount = 0

foreach ($lsoa in $pipLSOAs) {
    $code = $lsoa.LSOA21CD
    if ($hlnmByCode.ContainsKey($code)) {
        $h = $hlnmByCode[$code]

        # Determine mission classification from Economic Growth percentile
        $gp = [int]$h.'Economic Growth Percentile'
        $mission = if ($gp -ge 80) { "Mission Critical" }
                   elseif ($gp -ge 40) { "Mission Priority" }
                   elseif ($gp -ge 30) { "Mission Support" }
                   else { "Other" }

        $hlnmJS += "`"$code`"`:{"
        $hlnmJS += "gp:$($h.'Economic Growth Percentile'),"
        $hlnmJS += "ep:$($h.'Clean Energy Percentile'),"
        $hlnmJS += "cp:$($h.'Safe Streets Percentile'),"
        $hlnmJS += "opp:$($h.'Opportunity Percentile'),"
        $hlnmJS += "hp:$($h.'Health Percentile'),"
        $hlnmJS += "or:$($h.'Overall Rank'),"
        $hlnmJS += "t:`"$mission`","
        $hlnmJS += "n:`"$($lsoa.LSOA21NM -replace '"', '\"')`""
        $hlnmJS += "},"
        $hlnmCount++
    }
}
$hlnmJS = $hlnmJS.TrimEnd(',') + "};"
$hlnmJS | Out-File "lsoa_hlnm_data_complete.js" -Encoding utf8
Write-Host "  Generated lsoa_hlnm_data_complete.js with $hlnmCount LSOAs"

# Build Economic JS file
$econJS = "const LSOA_ECONOMIC_DATA = {"
$econCount = 0

foreach ($lsoa in $pipLSOAs) {
    $code = $lsoa.LSOA21CD
    if ($econByCode.ContainsKey($code)) {
        $e = $econByCode[$code]

        $econJS += "`"$code`"`:{"
        $econJS += "jsa:$($e.JSA),"
        $econJS += "uc_search:$($e.UC_searching),"
        $econJS += "uc_total:$($e.UC_total),"
        $econJS += "jobs_density:$($e.Jobs_Density),"
        $econJS += "jobs_access:$($e.Jobs_Accessibility),"
        $econJS += "income:$($e.Median_Income),"
        $econJS += "no_quals:$($e.No_Quals),"
        $econJS += "level3plus:$($e.Level_3_plus),"
        $econJS += "higher_mgr:$($e.Higher_Mgr_Prof),"
        $econJS += "digital:$($e.Digital),"
        $econJS += "broadband:$($e.Broadband),"
        $econJS += "highgrowth:$($e.High_Growth)"
        $econJS += "},"
        $econCount++
    }
}
$econJS = $econJS.TrimEnd(',') + "};"
$econJS | Out-File "lsoa_economic_data_complete.js" -Encoding utf8
Write-Host "  Generated lsoa_economic_data_complete.js with $econCount LSOAs"

# Build Map JS file (grouped by MSOA)
$mapJS = "const LSOA_MAP_DATA = {`n"
$mapCount = 0

$lsoasByMSOA = $pipLSOAs | Group-Object -Property MSOA21CD | Sort-Object Name

foreach ($msoaGroup in $lsoasByMSOA) {
    $msoaCode = $msoaGroup.Name
    $mapJS += "`"$msoaCode`":[`n"

    foreach ($lsoa in $msoaGroup.Group) {
        $code = $lsoa.LSOA21CD
        $name = $lsoa.LSOA21NM -replace '"', '\"'

        # Get mission from HLNM data
        $mission = "Other"
        if ($hlnmByCode.ContainsKey($code)) {
            $gp = [int]$hlnmByCode[$code].'Economic Growth Percentile'
            $mission = if ($gp -ge 80) { "Mission Critical" }
                       elseif ($gp -ge 40) { "Mission Priority" }
                       elseif ($gp -ge 30) { "Mission Support" }
                       else { "Other" }
        }

        # Get coordinates
        $lat = 52.5
        $lng = -1.5
        if ($coordsByCode.ContainsKey($code)) {
            $lat = $coordsByCode[$code].lat
            $lng = $coordsByCode[$code].lng
        }

        $mapJS += "{c:`"$code`",n:`"$name`",lat:$lat,lng:$lng,m:`"$mission`"},`n"
        $mapCount++
    }

    $mapJS = $mapJS.TrimEnd(",`n") + "`n],`n"
}

$mapJS = $mapJS.TrimEnd(",`n") + "`n};"
$mapJS | Out-File "lsoa_map_data_complete.js" -Encoding utf8
Write-Host "  Generated lsoa_map_data_complete.js with $mapCount LSOAs"

Write-Host "`n=== Summary ===" -ForegroundColor Green
Write-Host "  Total LSOAs processed: $($pipLSOAs.Count)"
Write-Host "  HLNM data: $hlnmCount LSOAs"
Write-Host "  Economic data: $econCount LSOAs"
Write-Host "  Map data: $mapCount LSOAs"
Write-Host "`nNext steps:"
Write-Host "  1. Review the generated files"
Write-Host "  2. Replace old JS files with new ones"
Write-Host "  3. Test in browser`n"
