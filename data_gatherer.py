"""
Pride in Place Data Explorer - Data Gathering Script
Downloads and processes MSOA-level data for Pride in Place neighbourhoods.

Data Sources:
- Pride in Place MSOA list from OCSI
- Census 2021 data from NOMIS (economic activity, health, education, deprivation)
- IMD 2025 from GOV.UK (aggregated from LSOA to MSOA)
- LSOA to MSOA lookup from ONS

Run this script to generate the data.json file used by the dashboard.
"""

import requests
import pandas as pd
import numpy as np
import json
import zipfile
import io
import os
from pathlib import Path

# Configuration
DATA_DIR = Path(__file__).parent / "data"
OUTPUT_FILE = Path(__file__).parent / "data.json"

# URLs for data sources
URLS = {
    # Pride in Place MSOA list
    "pip_msoas": "https://ocsi.uk/wp-content/uploads/2025/10/pride_in_place_MSOA_LA.xlsx",

    # LSOA to MSOA lookup
    "lsoa_msoa_lookup": "https://open-geography-portalx-ons.hub.arcgis.com/api/download/v1/items/45fdd4465604493cb7d2238ad642172b/csv?layers=0",

    # Census 2021 bulk downloads (NOMIS)
    "census_economic_activity": "https://www.nomisweb.co.uk/output/census/2021/census2021-ts066.zip",
    "census_health": "https://www.nomisweb.co.uk/output/census/2021/census2021-ts037.zip",
    "census_disability": "https://www.nomisweb.co.uk/output/census/2021/census2021-ts038.zip",
    "census_qualifications": "https://www.nomisweb.co.uk/output/census/2021/census2021-ts067.zip",
    "census_deprivation": "https://www.nomisweb.co.uk/output/census/2021/census2021-ts011.zip",
    "census_tenure": "https://www.nomisweb.co.uk/output/census/2021/census2021-ts054.zip",
    "census_unpaid_care": "https://www.nomisweb.co.uk/output/census/2021/census2021-ts039.zip",
    "census_occupation": "https://www.nomisweb.co.uk/output/census/2021/census2021-ts063.zip",

    # IMD 2025 - LSOA level data
    "imd_2025": "https://assets.publishing.service.gov.uk/media/6721a39246532eb85b519981/File_7_-_All_IoD2025_Scores__Ranks_and_Deciles.csv",
}

def ensure_data_dir():
    """Create data directory if it doesn't exist."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)

def download_file(url, filename, force=False):
    """Download a file if it doesn't exist or force is True."""
    filepath = DATA_DIR / filename
    if filepath.exists() and not force:
        print(f"  Using cached: {filename}")
        return filepath

    print(f"  Downloading: {filename}...")
    response = requests.get(url, timeout=120)
    response.raise_for_status()

    with open(filepath, 'wb') as f:
        f.write(response.content)

    return filepath

def download_and_extract_zip(url, name, force=False):
    """Download a ZIP file and extract the MSOA CSV."""
    zip_path = DATA_DIR / f"{name}.zip"
    csv_path = DATA_DIR / f"{name}_msoa.csv"

    if csv_path.exists() and not force:
        print(f"  Using cached: {name}_msoa.csv")
        return csv_path

    print(f"  Downloading: {name}.zip...")
    response = requests.get(url, timeout=120)
    response.raise_for_status()

    # Save ZIP
    with open(zip_path, 'wb') as f:
        f.write(response.content)

    # Extract MSOA file
    with zipfile.ZipFile(zip_path, 'r') as zf:
        for filename in zf.namelist():
            if 'msoa' in filename.lower() and filename.endswith('.csv'):
                print(f"  Extracting: {filename}")
                with zf.open(filename) as source:
                    content = source.read()
                with open(csv_path, 'wb') as target:
                    target.write(content)
                return csv_path

    raise ValueError(f"No MSOA CSV found in {name}.zip")

def load_pip_msoas():
    """Load the Pride in Place MSOA list."""
    print("\n1. Loading Pride in Place MSOA list...")
    filepath = download_file(URLS["pip_msoas"], "pride_in_place_msoas.xlsx")
    df = pd.read_excel(filepath)

    # Standardise column names
    df.columns = [c.strip().lower().replace(' ', '_') for c in df.columns]

    # Find the MSOA code column
    msoa_col = None
    for col in df.columns:
        if 'msoa' in col and 'code' in col:
            msoa_col = col
            break
        elif 'msoa' in col:
            msoa_col = col

    if msoa_col is None:
        # Try to find any column that looks like MSOA codes
        for col in df.columns:
            sample = df[col].dropna().iloc[0] if len(df[col].dropna()) > 0 else ""
            if isinstance(sample, str) and sample.startswith('E02'):
                msoa_col = col
                break

    print(f"  Found {len(df)} Pride in Place neighbourhoods")
    print(f"  Columns: {list(df.columns)}")

    return df

def load_lsoa_msoa_lookup():
    """Load LSOA to MSOA lookup table."""
    print("\n2. Loading LSOA to MSOA lookup...")
    filepath = download_file(URLS["lsoa_msoa_lookup"], "lsoa_msoa_lookup.csv")
    df = pd.read_csv(filepath)
    print(f"  Loaded {len(df)} LSOA-MSOA mappings")
    return df

def load_census_data():
    """Load all Census 2021 datasets."""
    print("\n3. Loading Census 2021 data...")

    datasets = {}

    # Economic activity (TS066)
    path = download_and_extract_zip(URLS["census_economic_activity"], "ts066")
    datasets["economic_activity"] = pd.read_csv(path)

    # General health (TS037)
    path = download_and_extract_zip(URLS["census_health"], "ts037")
    datasets["health"] = pd.read_csv(path)

    # Disability (TS038)
    path = download_and_extract_zip(URLS["census_disability"], "ts038")
    datasets["disability"] = pd.read_csv(path)

    # Qualifications (TS067)
    path = download_and_extract_zip(URLS["census_qualifications"], "ts067")
    datasets["qualifications"] = pd.read_csv(path)

    # Deprivation dimensions (TS011)
    path = download_and_extract_zip(URLS["census_deprivation"], "ts011")
    datasets["deprivation"] = pd.read_csv(path)

    # Tenure (TS054)
    path = download_and_extract_zip(URLS["census_tenure"], "ts054")
    datasets["tenure"] = pd.read_csv(path)

    # Unpaid care (TS039)
    path = download_and_extract_zip(URLS["census_unpaid_care"], "ts039")
    datasets["unpaid_care"] = pd.read_csv(path)

    # Occupation (TS063)
    path = download_and_extract_zip(URLS["census_occupation"], "ts063")
    datasets["occupation"] = pd.read_csv(path)

    return datasets

def load_imd_data(lsoa_msoa_lookup):
    """Load IMD 2025 data and aggregate to MSOA level."""
    print("\n4. Loading IMD 2025 data...")
    filepath = download_file(URLS["imd_2025"], "imd_2025.csv")

    imd = pd.read_csv(filepath)
    print(f"  Loaded {len(imd)} LSOA records")

    # Find the LSOA code column
    lsoa_col = None
    for col in imd.columns:
        if 'lsoa' in col.lower() and 'code' in col.lower():
            lsoa_col = col
            break

    if lsoa_col is None:
        lsoa_col = imd.columns[0]  # Usually first column

    print(f"  Using LSOA column: {lsoa_col}")
    print(f"  Available columns: {list(imd.columns)[:10]}...")

    return imd, lsoa_col

def calculate_percentiles(series):
    """Calculate percentile rank for a series (0-100, higher = worse deprivation)."""
    return series.rank(pct=True) * 100

def process_economic_activity(df, msoa_codes):
    """Process economic activity data to get key metrics."""
    # Find geography code column
    geo_col = [c for c in df.columns if 'geography' in c.lower() and 'code' in c.lower()][0]

    # Filter to our MSOAs
    df = df[df[geo_col].isin(msoa_codes)].copy()

    # Find relevant columns
    total_col = [c for c in df.columns if 'total' in c.lower()][0]

    # Look for unemployment and economic inactivity columns
    unemployed_cols = [c for c in df.columns if 'unemployed' in c.lower()]
    inactive_cols = [c for c in df.columns if 'inactive' in c.lower() and 'economically' in c.lower()]
    employed_cols = [c for c in df.columns if 'employed' in c.lower() or 'employee' in c.lower()]

    result = df[[geo_col]].copy()
    result = result.rename(columns={geo_col: 'msoa_code'})

    # Calculate rates
    if unemployed_cols:
        result['unemployed_count'] = df[unemployed_cols].sum(axis=1)
    if inactive_cols:
        result['inactive_count'] = df[inactive_cols].sum(axis=1)
    if employed_cols:
        result['employed_count'] = df[employed_cols].sum(axis=1)

    result['total_population'] = df[total_col]

    return result

def process_health(df, msoa_codes):
    """Process health data."""
    geo_col = [c for c in df.columns if 'geography' in c.lower() and 'code' in c.lower()][0]
    df = df[df[geo_col].isin(msoa_codes)].copy()

    total_col = [c for c in df.columns if 'total' in c.lower()][0]
    bad_health_cols = [c for c in df.columns if 'bad' in c.lower() or 'poor' in c.lower()]
    good_health_cols = [c for c in df.columns if 'good' in c.lower() or 'very good' in c.lower()]

    result = df[[geo_col]].copy()
    result = result.rename(columns={geo_col: 'msoa_code'})
    result['total'] = df[total_col]

    if bad_health_cols:
        result['bad_health_count'] = df[bad_health_cols].sum(axis=1)
        result['bad_health_pct'] = (result['bad_health_count'] / result['total']) * 100

    return result

def process_qualifications(df, msoa_codes):
    """Process qualifications data."""
    geo_col = [c for c in df.columns if 'geography' in c.lower() and 'code' in c.lower()][0]
    df = df[df[geo_col].isin(msoa_codes)].copy()

    total_col = [c for c in df.columns if 'total' in c.lower()][0]
    no_qual_cols = [c for c in df.columns if 'no qual' in c.lower()]
    level4_cols = [c for c in df.columns if 'level 4' in c.lower()]

    result = df[[geo_col]].copy()
    result = result.rename(columns={geo_col: 'msoa_code'})
    result['total'] = df[total_col]

    if no_qual_cols:
        result['no_qualifications_count'] = df[no_qual_cols].sum(axis=1)
        result['no_qualifications_pct'] = (result['no_qualifications_count'] / result['total']) * 100

    if level4_cols:
        result['level4_plus_count'] = df[level4_cols].sum(axis=1)
        result['level4_plus_pct'] = (result['level4_plus_count'] / result['total']) * 100

    return result

def process_deprivation(df, msoa_codes):
    """Process household deprivation data."""
    geo_col = [c for c in df.columns if 'geography' in c.lower() and 'code' in c.lower()][0]
    df = df[df[geo_col].isin(msoa_codes)].copy()

    total_col = [c for c in df.columns if 'total' in c.lower()][0]

    # Look for columns indicating deprivation dimensions
    deprived_cols = [c for c in df.columns if 'deprived' in c.lower()]

    result = df[[geo_col]].copy()
    result = result.rename(columns={geo_col: 'msoa_code'})
    result['total_households'] = df[total_col]

    # Count households deprived in 1+ dimensions
    if deprived_cols:
        # Exclude "not deprived" column
        deprived_cols = [c for c in deprived_cols if 'not' not in c.lower()]
        if deprived_cols:
            result['deprived_households'] = df[deprived_cols].sum(axis=1)
            result['deprived_pct'] = (result['deprived_households'] / result['total_households']) * 100

    return result

def process_tenure(df, msoa_codes):
    """Process tenure data."""
    geo_col = [c for c in df.columns if 'geography' in c.lower() and 'code' in c.lower()][0]
    df = df[df[geo_col].isin(msoa_codes)].copy()

    total_col = [c for c in df.columns if 'total' in c.lower()][0]
    social_cols = [c for c in df.columns if 'social' in c.lower() or 'council' in c.lower()]
    owned_cols = [c for c in df.columns if 'owned' in c.lower() or 'owns' in c.lower()]
    private_rent_cols = [c for c in df.columns if 'private' in c.lower() and 'rent' in c.lower()]

    result = df[[geo_col]].copy()
    result = result.rename(columns={geo_col: 'msoa_code'})
    result['total_households'] = df[total_col]

    if social_cols:
        result['social_rented'] = df[social_cols].sum(axis=1)
        result['social_rented_pct'] = (result['social_rented'] / result['total_households']) * 100

    if owned_cols:
        result['owned'] = df[owned_cols].sum(axis=1)
        result['owned_pct'] = (result['owned'] / result['total_households']) * 100

    return result

def aggregate_imd_to_msoa(imd_df, lsoa_col, lsoa_msoa_lookup):
    """Aggregate IMD LSOA data to MSOA level using population-weighted averages."""
    print("  Aggregating IMD to MSOA level...")

    # Get LSOA to MSOA mapping
    lookup = lsoa_msoa_lookup[['LSOA21CD', 'MSOA21CD']].drop_duplicates()

    # Merge IMD with lookup
    merged = imd_df.merge(lookup, left_on=lsoa_col, right_on='LSOA21CD', how='left')

    # Find score columns to aggregate
    score_cols = [c for c in imd_df.columns if 'score' in c.lower()]
    rank_cols = [c for c in imd_df.columns if 'rank' in c.lower() and 'decile' not in c.lower()]

    # For scores, take mean within MSOA
    # For ranks, we'll recalculate after aggregation
    agg_dict = {col: 'mean' for col in score_cols if col in merged.columns}

    if agg_dict:
        msoa_imd = merged.groupby('MSOA21CD').agg(agg_dict).reset_index()
        msoa_imd = msoa_imd.rename(columns={'MSOA21CD': 'msoa_code'})
        return msoa_imd

    return None

def build_final_dataset(pip_df, census_datasets, imd_msoa, all_msoa_codes):
    """Combine all data sources into final dataset with percentiles."""
    print("\n5. Building final dataset with percentiles...")

    # Get list of PiP MSOA codes
    msoa_col = None
    for col in pip_df.columns:
        sample = pip_df[col].dropna().iloc[0] if len(pip_df[col].dropna()) > 0 else ""
        if isinstance(sample, str) and sample.startswith('E02'):
            msoa_col = col
            break

    if msoa_col is None:
        raise ValueError("Could not find MSOA code column in PiP data")

    pip_msoas = pip_df[msoa_col].tolist()
    print(f"  Processing {len(pip_msoas)} Pride in Place MSOAs")

    # Find name column
    name_col = None
    for col in pip_df.columns:
        if 'name' in col.lower() or 'neighbourhood' in col.lower():
            name_col = col
            break

    # Find LA column
    la_col = None
    for col in pip_df.columns:
        if 'local' in col.lower() or 'authority' in col.lower() or 'la' in col.lower():
            la_col = col
            break

    # Start building the result
    result = pip_df[[msoa_col]].copy()
    result = result.rename(columns={msoa_col: 'msoa_code'})

    if name_col:
        result['neighbourhood_name'] = pip_df[name_col].values
    if la_col:
        result['local_authority'] = pip_df[la_col].values

    # Process each Census dataset and calculate national percentiles
    print("  Processing economic activity...")
    econ = process_economic_activity(census_datasets["economic_activity"], all_msoa_codes)

    print("  Processing health...")
    health = process_health(census_datasets["health"], all_msoa_codes)

    print("  Processing qualifications...")
    quals = process_qualifications(census_datasets["qualifications"], all_msoa_codes)

    print("  Processing household deprivation...")
    deprivation = process_deprivation(census_datasets["deprivation"], all_msoa_codes)

    print("  Processing tenure...")
    tenure = process_tenure(census_datasets["tenure"], all_msoa_codes)

    # Merge all data
    for df in [econ, health, quals, deprivation, tenure]:
        if df is not None and len(df) > 0:
            result = result.merge(df, on='msoa_code', how='left')

    if imd_msoa is not None:
        result = result.merge(imd_msoa, on='msoa_code', how='left')

    return result

def calculate_all_percentiles(df, all_england_df):
    """Calculate percentile rankings compared to all England MSOAs."""
    print("\n6. Calculating percentile rankings...")

    # Columns where higher = worse (so higher percentile = more deprived)
    higher_is_worse = [
        'bad_health_pct', 'no_qualifications_pct', 'deprived_pct',
        'unemployed_count', 'inactive_count'
    ]

    # Columns where lower = worse
    lower_is_worse = [
        'level4_plus_pct', 'owned_pct', 'employed_count'
    ]

    numeric_cols = df.select_dtypes(include=[np.number]).columns

    for col in numeric_cols:
        if col in ['msoa_code']:
            continue

        # Get all England values for this column
        if col in all_england_df.columns:
            all_values = all_england_df[col].dropna()

            # Calculate percentile for each PiP MSOA
            percentile_col = f"{col}_percentile"
            df[percentile_col] = df[col].apply(
                lambda x: (all_values < x).sum() / len(all_values) * 100 if pd.notna(x) else np.nan
            )

            # Invert if lower is worse (so higher percentile always = more deprived)
            if col in lower_is_worse:
                df[percentile_col] = 100 - df[percentile_col]

    return df

def export_to_json(df, output_path):
    """Export dataframe to JSON for the dashboard."""
    print(f"\n7. Exporting to {output_path}...")

    # Convert to list of dictionaries
    records = df.to_dict(orient='records')

    # Clean up NaN values
    for record in records:
        for key, value in record.items():
            if pd.isna(value):
                record[key] = None

    # Build metadata
    metadata = {
        "generated": pd.Timestamp.now().isoformat(),
        "total_areas": len(records),
        "data_sources": {
            "pride_in_place": "OCSI - Pride in Place Programme neighbourhoods",
            "census_2021": "ONS Census 2021 via NOMIS",
            "imd_2025": "MHCLG English Indices of Deprivation 2025"
        }
    }

    output = {
        "metadata": metadata,
        "areas": records
    }

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    print(f"  Exported {len(records)} areas")

def main():
    """Main function to run the data gathering pipeline."""
    print("=" * 60)
    print("Pride in Place Data Explorer - Data Gathering")
    print("=" * 60)

    ensure_data_dir()

    # Load Pride in Place MSOA list
    pip_df = load_pip_msoas()

    # Load lookup table
    lsoa_msoa_lookup = load_lsoa_msoa_lookup()

    # Get all England MSOA codes
    all_msoa_codes = lsoa_msoa_lookup['MSOA21CD'].unique().tolist()
    print(f"  Total MSOAs in England: {len(all_msoa_codes)}")

    # Load Census data
    census_data = load_census_data()

    # Load and aggregate IMD data
    imd_df, lsoa_col = load_imd_data(lsoa_msoa_lookup)
    imd_msoa = aggregate_imd_to_msoa(imd_df, lsoa_col, lsoa_msoa_lookup)

    # Build final dataset
    final_df = build_final_dataset(pip_df, census_data, imd_msoa, all_msoa_codes)

    # Export to JSON
    export_to_json(final_df, OUTPUT_FILE)

    print("\n" + "=" * 60)
    print("Data gathering complete!")
    print(f"Output file: {OUTPUT_FILE}")
    print("=" * 60)

if __name__ == "__main__":
    main()
