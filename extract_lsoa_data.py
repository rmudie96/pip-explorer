import csv
import json

# Read the embedded DATA from index.html to get the 40 MSOA codes
msoa_codes = set([
    "E02001954", "E02001947", "E02001948", "E02001949", "E02001950",
    "E02001951", "E02001952", "E02001953", "E02001405", "E02001406",
    "E02001407", "E02001408", "E02001094", "E02001095", "E02001096",
    "E02001097", "E02002385", "E02002386", "E02002387", "E02002388",
    "E02002651", "E02002652", "E02002653", "E02001305", "E02001306",
    "E02001307", "E02004205", "E02004206", "E02004207", "E02005355",
    "E02005356", "E02005357", "E02005440", "E02005441", "E02005442",
    "E02006855", "E02006502", "E02003505", "E02003506", "E02006545"
])

# Read LSOA to MSOA lookup
lsoa_to_msoa = {}
with open('lsoa_msoa_la_region_lookup_csv.csv', 'r', encoding='utf-8-sig') as f:
    reader = csv.DictReader(f)
    for row in reader:
        if row['MSOA21CD'] in msoa_codes:
            lsoa_to_msoa[row['LSOA21CD']] = row['MSOA21CD']

print(f"Found {len(lsoa_to_msoa)} LSOAs across 40 MSOAs")

# Read HLNM classifications
lsoa_classifications = {}
with open('Hyper-Local-Need-Measure-2025-_csv.csv', 'r', encoding='utf-8-sig') as f:
    lines = f.readlines()
    for line in lines[5:]:  # Skip header rows
        parts = line.strip().split(',')
        if len(parts) > 9:
            lsoa_code = parts[0].strip()
            typology = parts[9].strip()
            if lsoa_code in lsoa_to_msoa:
                lsoa_classifications[lsoa_code] = typology

print(f"Found {len(lsoa_classifications)} LSOA classifications")

# Read LSOA centroids
lsoa_centroids = {}
with open('Lower_layer_Super_Output_Areas_December_2021_Boundaries_EW_BSC_V4_3901388190129020682 (1).csv', 'r', encoding='utf-8-sig') as f:
    reader = csv.DictReader(f)
    for row in reader:
        if row['LSOA21CD'] in lsoa_to_msoa:
            lsoa_centroids[row['LSOA21CD']] = {
                'name': row['LSOA21NM'],
                'lat': float(row['LAT']),
                'lng': float(row['LONG'])
            }

print(f"Found {len(lsoa_centroids)} LSOA centroids")

# Build final structure grouped by MSOA
msoa_lsoa_data = {}
for lsoa_code, msoa_code in lsoa_to_msoa.items():
    if msoa_code not in msoa_lsoa_data:
        msoa_lsoa_data[msoa_code] = []
    
    if lsoa_code in lsoa_centroids:
        msoa_lsoa_data[msoa_code].append({
            'code': lsoa_code,
            'name': lsoa_centroids[lsoa_code]['name'],
            'lat': lsoa_centroids[lsoa_code]['lat'],
            'lng': lsoa_centroids[lsoa_code]['lng'],
            'classification': lsoa_classifications.get(lsoa_code, '')
        })

# Write to JavaScript file
with open('lsoa_embedded_data.js', 'w') as f:
    f.write('const LSOA_DATA = ')
    json.dump(msoa_lsoa_data, f, indent=2)
    f.write(';')

print(f"Written LSOA data for {len(msoa_lsoa_data)} MSOAs to lsoa_embedded_data.js")
