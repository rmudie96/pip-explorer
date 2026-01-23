# Process HLNM data and update data.json
# This script reads HLNM_MSOA_csv.csv and adds HLNM indicators to data.json

$projectPath = "C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer"

Write-Host "Processing HLNM data..." -ForegroundColor Cyan

# Read the existing data.json
$dataJsonPath = Join-Path $projectPath "data.json"
$dataJson = Get-Content $dataJsonPath -Raw | ConvertFrom-Json

# Read the HLNM MSOA CSV
$hlnmPath = Join-Path $projectPath "HLNM_MSOA_csv.csv"
$hlnmData = Import-Csv $hlnmPath

Write-Host "Loaded $($hlnmData.Count) MSOAs from HLNM data" -ForegroundColor Green
Write-Host "PiP areas to process: $($dataJson.areas.Count)" -ForegroundColor Green

# Create lookup dictionary for HLNM data
$hlnmLookup = @{}
foreach ($row in $hlnmData) {
    $code = $row.'Area Code'
    if ($code) {
        $hlnmLookup[$code] = @{
            Growth = [double]$row.Growth
            Energy = [double]$row.Energy
            Crime = [int]$row.Crime
            Opportunity = [double]$row.Opportunity
            Health = [double]$row.Health
        }
    }
}

Write-Host "Created HLNM lookup with $($hlnmLookup.Count) entries" -ForegroundColor Green

# Calculate percentiles for each indicator
# Higher score = higher need (worse) for Growth, Energy, Opportunity, Health
# For Crime: the value is a rank where lower = worse (higher crime)

$allGrowth = $hlnmData | ForEach-Object { [double]$_.Growth } | Sort-Object
$allEnergy = $hlnmData | ForEach-Object { [double]$_.Energy } | Sort-Object
$allCrime = $hlnmData | ForEach-Object { [int]$_.Crime } | Sort-Object
$allOpportunity = $hlnmData | ForEach-Object { [double]$_.Opportunity } | Sort-Object
$allHealth = $hlnmData | ForEach-Object { [double]$_.Health } | Sort-Object

$totalMSOAs = $hlnmData.Count

function Get-Percentile {
    param (
        [double]$value,
        [array]$sortedValues,
        [bool]$lowerIsWorse = $false
    )

    $count = ($sortedValues | Where-Object { $_ -le $value }).Count
    $percentile = [math]::Round(($count / $sortedValues.Count) * 100, 1)

    if ($lowerIsWorse) {
        # Invert percentile for Crime where lower rank = worse
        $percentile = 100 - $percentile
    }

    return $percentile
}

# Process each PiP area
$matchedCount = 0
$notFoundCodes = @()

foreach ($area in $dataJson.areas) {
    $msoa = $area.msoa_code

    if ($hlnmLookup.ContainsKey($msoa)) {
        $hlnm = $hlnmLookup[$msoa]
        $matchedCount++

        # Add HLNM scores
        $area | Add-Member -NotePropertyName "hlnm_growth" -NotePropertyValue $hlnm.Growth -Force
        $area | Add-Member -NotePropertyName "hlnm_energy" -NotePropertyValue $hlnm.Energy -Force
        $area | Add-Member -NotePropertyName "hlnm_crime" -NotePropertyValue $hlnm.Crime -Force
        $area | Add-Member -NotePropertyName "hlnm_opportunity" -NotePropertyValue $hlnm.Opportunity -Force
        $area | Add-Member -NotePropertyName "hlnm_health" -NotePropertyValue $hlnm.Health -Force

        # Calculate percentiles
        # For Growth, Energy, Opportunity, Health: higher score = higher need (worse)
        # Percentile shows what % of areas have LOWER need than this area
        $growthPct = Get-Percentile -value $hlnm.Growth -sortedValues $allGrowth
        $energyPct = Get-Percentile -value $hlnm.Energy -sortedValues $allEnergy
        $opportunityPct = Get-Percentile -value $hlnm.Opportunity -sortedValues $allOpportunity
        $healthPct = Get-Percentile -value $hlnm.Health -sortedValues $allHealth

        # For Crime: lower rank number = higher crime/worse
        # We want percentile to show what % have lower crime (better)
        $crimePct = Get-Percentile -value $hlnm.Crime -sortedValues $allCrime -lowerIsWorse $true

        $area | Add-Member -NotePropertyName "hlnm_growth_percentile" -NotePropertyValue $growthPct -Force
        $area | Add-Member -NotePropertyName "hlnm_energy_percentile" -NotePropertyValue $energyPct -Force
        $area | Add-Member -NotePropertyName "hlnm_crime_percentile" -NotePropertyValue $crimePct -Force
        $area | Add-Member -NotePropertyName "hlnm_opportunity_percentile" -NotePropertyValue $opportunityPct -Force
        $area | Add-Member -NotePropertyName "hlnm_health_percentile" -NotePropertyValue $healthPct -Force

        Write-Host "  Matched: $msoa ($($area.neighbourhood_name))" -ForegroundColor Gray
    } else {
        $notFoundCodes += $msoa
        Write-Host "  NOT FOUND: $msoa ($($area.neighbourhood_name))" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  Matched: $matchedCount / $($dataJson.areas.Count) PiP areas" -ForegroundColor Green

if ($notFoundCodes.Count -gt 0) {
    Write-Host "  Not found: $($notFoundCodes.Count) areas" -ForegroundColor Yellow
    Write-Host "  Missing codes: $($notFoundCodes -join ', ')" -ForegroundColor Yellow
}

# Update metadata
$dataJson.metadata | Add-Member -NotePropertyName "hlnm_source" -NotePropertyValue "OCSI Hyper-Local Need Measure 2024" -Force
$dataJson.metadata | Add-Member -NotePropertyName "hlnm_added" -NotePropertyValue (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ") -Force
$dataJson.metadata | Add-Member -NotePropertyName "total_msoas_for_percentile" -NotePropertyValue $totalMSOAs -Force

# Save updated data.json
$outputPath = Join-Path $projectPath "data.json"
$dataJson | ConvertTo-Json -Depth 10 | Set-Content $outputPath -Encoding UTF8

Write-Host ""
Write-Host "Saved updated data to: $outputPath" -ForegroundColor Green
Write-Host "Done!" -ForegroundColor Cyan
