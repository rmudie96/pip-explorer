# Generate embedded data for HTML
$projectPath = "C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer"
$data = Get-Content (Join-Path $projectPath "data.json") -Raw | ConvertFrom-Json

# Build array for output
$areasJson = @()

foreach ($area in $data.areas) {
    $obj = [ordered]@{
        msoa_code = $area.msoa_code
        neighbourhood_name = $area.neighbourhood_name
        local_authority = $area.local_authority
        bad_health_pct = $area.bad_health_pct
        bad_health_pct_percentile = $area.bad_health_pct_percentile
        no_qualifications_pct = $area.no_qualifications_pct
        no_qualifications_pct_percentile = $area.no_qualifications_pct_percentile
        level4_plus_pct = $area.level4_plus_pct
        level4_plus_pct_percentile = $area.level4_plus_pct_percentile
        unemployed_count_percentile = $area.unemployed_count_percentile
        inactive_count_percentile = $area.inactive_count_percentile
        employed_count_percentile = $area.employed_count_percentile
        deprived_pct = $area.deprived_pct
        deprived_pct_percentile = $area.deprived_pct_percentile
        owned_pct = $area.owned_pct
        owned_pct_percentile = $area.owned_pct_percentile
        social_rented_pct = $area.social_rented_pct
        social_rented_pct_percentile = $area.social_rented_pct_percentile
    }

    # Add IMD fields with proper names
    $obj["Index of Multiple Deprivation (IMD) Score"] = $area."Index of Multiple Deprivation (IMD) Score"
    $obj["Income Score (rate)"] = $area."Income Score (rate)"
    $obj["Employment Score (rate)"] = $area."Employment Score (rate)"
    $obj["Education, Skills and Training Score"] = $area."Education, Skills and Training Score"
    $obj["Health Deprivation and Disability Score"] = $area."Health Deprivation and Disability Score"
    $obj["Crime Score"] = $area."Crime Score"
    $obj["Barriers to Housing and Services Score"] = $area."Barriers to Housing and Services Score"
    $obj["Living Environment Score"] = $area."Living Environment Score"

    # Add HLNM fields if they exist
    if ($null -ne $area.hlnm_growth) {
        $obj.hlnm_growth = [math]::Round($area.hlnm_growth, 2)
        $obj.hlnm_energy = [math]::Round($area.hlnm_energy, 2)
        $obj.hlnm_crime = $area.hlnm_crime
        $obj.hlnm_opportunity = [math]::Round($area.hlnm_opportunity, 2)
        $obj.hlnm_health = [math]::Round($area.hlnm_health, 2)
        $obj.hlnm_growth_percentile = $area.hlnm_growth_percentile
        $obj.hlnm_energy_percentile = $area.hlnm_energy_percentile
        $obj.hlnm_crime_percentile = $area.hlnm_crime_percentile
        $obj.hlnm_opportunity_percentile = $area.hlnm_opportunity_percentile
        $obj.hlnm_health_percentile = $area.hlnm_health_percentile
    }

    # Add CNI fields
    if ($null -ne $area.civic_assets_rank) { $obj.civic_assets_rank = $area.civic_assets_rank }
    if ($null -ne $area.connectedness_rank) { $obj.connectedness_rank = $area.connectedness_rank }
    if ($null -ne $area.active_engaged_rank) { $obj.active_engaged_rank = $area.active_engaged_rank }

    # Add LSOA breakdown
    if ($null -ne $area.lsoa_total) {
        $obj.lsoa_total = $area.lsoa_total
        $obj.lsoa_critical = $area.lsoa_critical
        $obj.lsoa_priority = $area.lsoa_priority
        $obj.lsoa_support = $area.lsoa_support
    }

    $areasJson += $obj
}

$output = [ordered]@{
    metadata = [ordered]@{
        generated = "2026-01-21"
        total_areas = 40
        note = "Data for 40 Pride in Place neighbourhoods with HLNM indicators"
    }
    areas = $areasJson
}

$json = $output | ConvertTo-Json -Depth 10 -Compress
$json | Set-Content (Join-Path $projectPath "embedded_data.json") -Encoding UTF8
Write-Host "Generated embedded_data.json"
Write-Host "Length: $($json.Length) characters"
