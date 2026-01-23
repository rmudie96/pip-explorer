#!/bin/bash

# Get all LSOA codes from embedded data
grep -oP 'c:"E\d{8}"' index.html | sed 's/c:"//g' | sed 's/"//g' | sort -u > lsoa_codes.txt

echo "Extracting boundaries for $(wc -l < lsoa_codes.txt) LSOAs..."

# This will take the large GeoJSON and filter it to only our LSOAs
# Creating a much smaller file that can be loaded directly
python3 << 'PYTHON'
import json

# Read LSOA codes we need
with open('lsoa_codes.txt', 'r') as f:
    needed_codes = set(line.strip() for line in f)

print(f"Loading GeoJSON for {len(needed_codes)} LSOAs...")

# Read the full GeoJSON
with open('Lower_layer_Super_Output_Areas_December_2021_Boundaries_EW_BSC_V4_-4299016806856585929.geojson', 'r') as f:
    data = json.load(f)

# Filter to only features we need
filtered_features = [
    feature for feature in data['features']
    if feature['properties']['LSOA21CD'] in needed_codes
]

print(f"Extracted {len(filtered_features)} boundaries")

# Create new GeoJSON with only our features
filtered_data = {
    'type': 'FeatureCollection',
    'features': filtered_features
}

# Write to new file
with open('pride_in_place_lsoa_boundaries.geojson', 'w') as f:
    json.dump(filtered_data, f)

import os
size_mb = os.path.getsize('pride_in_place_lsoa_boundaries.gegeojson') / (1024 * 1024)
print(f"Created pride_in_place_lsoa_boundaries.geojson ({size_mb:.1f}MB)")
PYTHON

