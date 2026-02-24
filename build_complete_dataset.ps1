# Build complete dataset for all 40 Pride in Place MSOAs and their LSOAs
Write-Host "Building complete Pride in Place dataset..." -ForegroundColor Cyan

# Get list of 40 Pride in Place MSOAs
$htmlContent = Get-Content "index.html" -Raw
$dataStart = $htmlContent.IndexOf('const DATA = ') + 13
$dataEnd = $htmlContent.IndexOf('};', $dataStart) + 1
$jsonStr = $htmlContent.Substring($dataStart, $dataEnd - $dataStart)
$data = $jsonStr | ConvertFrom-Json

$pipMSOAs = $data.areas | Select-Object -ExpandProperty msoa_code
Write-Host "Found $($pipMSOAs.Count) Pride in Place MSOAs`n"

# Load LSOA lookup
$lookup = Import-Csv "lsoa_msoa_la_region_lookup_csv.csv"
Write-Host "Loaded national LSOA lookup with $($lookup.Count) LSOAs`n"

# Get all LSOAs for PiP MSOAs
$pipLSOAs = $lookup | Where-Object { $pipMSOAs -contains $_.MSOA21CD }
Write-Host "Found $($pipLSOAs.Count) LSOAs across the 40 Pride in Place MSOAs`n"

# Group by MSOA
$lsoasByMSOA = $pipLSOAs | Group-Object -Property MSOA21CD

Write-Host "LSOAs per MSOA:"
Write-Host "=" * 50
foreach ($msoa in $data.areas | Sort-Object neighbourhood_name) {
    $lsoaCount = ($lsoasByMSOA | Where-Object { $_.Name -eq $msoa.msoa_code }).Count
    Write-Host "$($msoa.neighbourhood_name): $lsoaCount LSOAs"
}

Write-Host "`nTotal LSOAs to process: $($pipLSOAs.Count)"
Write-Host "`nThis script will generate:"
Write-Host "  1. Complete LSOA map data with lat/lng (needs GeoJSON processing)"
Write-Host "  2. LSOA HLNM data (from existing Hyper-Local-Need-Measure-2025.xlsx)"
Write-Host "  3. LSOA economic data (from existing Hyper-local Need Index_econ_underlying.csv)"
Write-Host "`nReady to build complete dataset."
