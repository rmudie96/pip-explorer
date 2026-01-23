# Extract LSOA-level HLNM data for Pride in Place neighbourhoods
# Converts ranks to percentiles where LOWER rank = WORSE (higher percentile)

$csvPath = "Hyper-Local-Need-Measure-2025-_csv.csv"
$lsoaMapPath = "lsoa_embedded_data_temp.js"
$outputPath = "lsoa_hlnm_data.js"

# Read LSOA map data to get list of all PiP LSOAs
$lsoaMapContent = Get-Content $lsoaMapPath -Raw
$pipLSOACodes = [regex]::Matches($lsoaMapContent, 'c:"(E\d+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique

Write-Host "Found $($pipLSOACodes.Count) unique LSOA codes in Pride in Place areas"

# Read HLNM CSV (skip first 5 header rows)
$hlnmData = Import-Csv $csvPath -Header @(
    'lsoa_code', 'lsoa_name', 'la_code', 'la_name',
    'overall_score', 'col6', 'col7', 'col8', 'col9', 'typology', 'col11',
    'overall_rank',
    'growth_score', 'growth_rank',
    'energy_score', 'energy_rank',
    'crime_score', 'crime_rank',
    'opportunity_score', 'opportunity_rank',
    'health_score', 'health_rank'
) | Select-Object -Skip 5

# Filter to only PiP LSOAs
$pipLSOAs = $hlnmData | Where-Object { $pipLSOACodes -contains $_.lsoa_code }

Write-Host "Found $($pipLSOAs.Count) LSOAs with HLNM data"

# Total LSOAs for percentile calculation
$totalLSOAs = 33754

# Create compact JSON structure
$lsoaData = @{}

foreach ($lsoa in $pipLSOAs) {
    # Convert ranks to percentiles: lower rank (1) = higher percentile (bad)
    # Percentile = (rank / total) * 100

    $lsoaData[$lsoa.lsoa_code] = @{
        n = $lsoa.lsoa_name
        t = $lsoa.typology
        or = [int]$lsoa.overall_rank
        op = [math]::Round(([int]$lsoa.overall_rank / $totalLSOAs) * 100)
        gr = [int]$lsoa.growth_rank
        gp = [math]::Round(([int]$lsoa.growth_rank / $totalLSOAs) * 100)
        er = [int]$lsoa.energy_rank
        ep = [math]::Round(([int]$lsoa.energy_rank / $totalLSOAs) * 100)
        cr = [int]$lsoa.crime_rank
        cp = [math]::Round(([int]$lsoa.crime_rank / $totalLSOAs) * 100)
        opr = [int]$lsoa.opportunity_rank
        opp = [math]::Round(([int]$lsoa.opportunity_rank / $totalLSOAs) * 100)
        hr = [int]$lsoa.health_rank
        hp = [math]::Round(([int]$lsoa.health_rank / $totalLSOAs) * 100)
    }
}

# Output as JavaScript
$json = $lsoaData | ConvertTo-Json -Compress -Depth 10
$output = "const LSOA_HLNM_DATA = $json;"

Set-Content -Path $outputPath -Value $output -Encoding UTF8

Write-Host "Created $outputPath with data for $($lsoaData.Count) LSOAs"
Write-Host "File size: $((Get-Item $outputPath).Length) bytes"
