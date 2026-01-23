import json

# Load LSOA mappings
with open('lsoa_embedded_data_temp.js', 'r', encoding='utf-8') as f:
    content = f.read()

# Extract all LSOA codes we need
needed_lsoas = set()
for line in content.split('\n'):
    if '{c:"' in line:
        parts = line.split('{c:"')
        for part in parts[1:]:
            lsoa_code = part.split('"')[0]
            needed_lsoas.add(lsoa_code)

print(f"Found {len(needed_lsoas)} unique LSOAs to extract")

# Load full GeoJSON
print("Loading full GeoJSON...")
with open('Lower_layer_Super_Output_Areas_December_2021_Boundaries_EW_BSC_V4_-4299016806856585929.geojson', 'r', encoding='utf-8') as f:
    full_geojson = json.load(f)

# Filter to only our LSOAs
print("Filtering features...")
filtered_features = [
    feature for feature in full_geojson['features']
    if feature['properties']['LSOA21CD'] in needed_lsoas
]

print(f"Extracted {len(filtered_features)} features")

# Create new GeoJSON
pip_geojson = {
    "type": "FeatureCollection",
    "features": filtered_features
}

# Save
print("Saving...")
with open('pip_lsoa_boundaries.geojson', 'w', encoding='utf-8') as f:
    json.dump(pip_geojson, f)

print("Done! Created pip_lsoa_boundaries.geojson")
