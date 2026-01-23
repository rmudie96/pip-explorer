# Extract underlying economic indicators for Pride in Place LSOAs
# Creates compact JSON with key indicators that explain economic growth scores

$csvPath = "Hyper-local Need Index_econ_underlying.csv"
$lsoaMapPath = "lsoa_embedded_data_temp.js"
$outputPath = "lsoa_economic_underlying.js"

# Read LSOA map data to get list of all PiP LSOAs
$lsoaMapContent = Get-Content $lsoaMapPath -Raw
$pipLSOACodes = [regex]::Matches($lsoaMapContent, 'c:"(E\d+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique

Write-Host "Found $($pipLSOACodes.Count) unique LSOA codes in Pride in Place areas"

# Read underlying CSV (skip first 5 header rows)
# Column mapping based on header analysis:
# 5: UC searching count, 6: UC searching rate
# 7: UC planning count, 8: UC planning rate
# 9: UC preparing count, 10: UC preparing rate
# 11: UC no work count, 12: UC no work rate
# 13: JSA count, 14: JSA rate
# 20: Jobs Density rate
# 21: GVA per head rate (2022)
# 22: GVA per head rate (2012)
# 23: % change GVA 2012-2022
# 24: High growth jobs count
# 25: High growth jobs rate
# 26: Median household income
# 27: Higher managerial count
# 28: Higher managerial rate
# 29: No quals count
# 30: No quals rate
# 31: Level 3+ count
# 32: Level 3+ rate
# 35: Broadband speed
# 36: Digital propensity
# 37: Jobs access score

$economicData = Import-Csv $csvPath -Header @(
    'lsoa_code', 'lsoa_name', 'la_code', 'la_name',
    'uc_search_cnt', 'uc_search_rate',
    'uc_plan_cnt', 'uc_plan_rate',
    'uc_prep_cnt', 'uc_prep_rate',
    'uc_nowork_cnt', 'uc_nowork_rate',
    'jsa_cnt', 'jsa_rate',
    'incap_cnt', 'incap_rate',
    'sda_cnt', 'sda_rate',
    'income_sup_cnt', 'income_sup_rate',
    'carers_cnt', 'carers_rate',
    'jobs_density',
    'gva_2022', 'gva_2012',
    'gva_change_pct',
    'highgrowth_cnt', 'highgrowth_rate',
    'median_income',
    'higher_mgr_cnt', 'higher_mgr_rate',
    'no_quals_cnt', 'no_quals_rate',
    'level3plus_cnt', 'level3plus_rate',
    'broadband_speed',
    'digital_prop',
    'jobs_access'
) | Select-Object -Skip 5

# Filter to only PiP LSOAs
$pipLSOAs = $economicData | Where-Object { $pipLSOACodes -contains $_.lsoa_code }

Write-Host "Found $($pipLSOAs.Count) LSOAs with economic data"

# Create compact JSON structure
$lsoaEconData = @{}

foreach ($lsoa in $pipLSOAs) {
    # Clean numeric values (remove quotes, commas, pound signs)
    $cleanNum = { param($val)
        if ([string]::IsNullOrWhiteSpace($val)) { return $null }
        $cleaned = $val -replace '[£,"\s]', ''
        $num = 0.0
        if ([double]::TryParse($cleaned, [ref]$num)) { return [math]::Round($num, 2) }
        return $null
    }

    $lsoaEconData[$lsoa.lsoa_code] = @{
        # Benefits & worklessness (rates as %)
        uc_search = & $cleanNum $lsoa.uc_search_rate
        uc_total = (& $cleanNum $lsoa.uc_search_rate) + (& $cleanNum $lsoa.uc_plan_rate) + (& $cleanNum $lsoa.uc_prep_rate)
        jsa = & $cleanNum $lsoa.jsa_rate

        # Economic productivity
        jobs_density = & $cleanNum $lsoa.jobs_density
        gva = & $cleanNum $lsoa.gva_2022
        gva_change = & $cleanNum $lsoa.gva_change_pct
        highgrowth = & $cleanNum $lsoa.highgrowth_rate

        # Income & skills
        income = & $cleanNum $lsoa.median_income
        higher_mgr = & $cleanNum $lsoa.higher_mgr_rate
        no_quals = & $cleanNum $lsoa.no_quals_rate
        level3plus = & $cleanNum $lsoa.level3plus_rate

        # Digital & connectivity
        broadband = & $cleanNum $lsoa.broadband_speed
        digital = & $cleanNum $lsoa.digital_prop
        jobs_access = & $cleanNum $lsoa.jobs_access
    }
}

# Output as JavaScript
$json = $lsoaEconData | ConvertTo-Json -Compress -Depth 10
$output = "const LSOA_ECONOMIC_DATA = $json;"

Set-Content -Path $outputPath -Value $output -Encoding UTF8

Write-Host "Created $outputPath with data for $($lsoaEconData.Count) LSOAs"
Write-Host "File size: $((Get-Item $outputPath).Length) bytes"
