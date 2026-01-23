#!/bin/bash

# All 40 MSOA codes
MSOA_CODES="E02001954 E02001947 E02001948 E02001949 E02001950 E02001951 E02001952 E02001953 E02001405 E02001406 E02001407 E02001408 E02001094 E02001095 E02001096 E02001097 E02002385 E02002386 E02002387 E02002388 E02002651 E02002652 E02002653 E02001305 E02001306 E02001307 E02004205 E02004206 E02004207 E02005355 E02005356 E02005357 E02005440 E02005441 E02005442 E02006855 E02006502 E02003505 E02003506 E02006545"

echo "const LSOA_MAP_DATA = {"

for MSOA in $MSOA_CODES; do
    echo "\"$MSOA\":["
    
    # Get LSOAs for this MSOA from lookup
    awk -F',' -v msoa="$MSOA" 'NR>1 && $4==msoa {print $1}' lsoa_msoa_la_region_lookup_csv.csv | while read LSOA; do
        # Get centroid
        LAT=$(awk -F',' -v lsoa="$LSOA" '$2==lsoa {print $7; exit}' "Lower_layer_Super_Output_Areas_December_2021_Boundaries_EW_BSC_V4_3901388190129020682 (1).csv")
        LNG=$(awk -F',' -v lsoa="$LSOA" '$2==lsoa {print $8; exit}' "Lower_layer_Super_Output_Areas_December_2021_Boundaries_EW_BSC_V4_3901388190129020682 (1).csv")
        NAME=$(awk -F',' -v lsoa="$LSOA" '$2==lsoa {print $3; exit}' "Lower_layer_Super_Output_Areas_December_2021_Boundaries_EW_BSC_V4_3901388190129020682 (1).csv")
        MISSION=$(awk -F',' -v lsoa="$LSOA" 'NR>5 && $1==lsoa {print $10; exit}' Hyper-Local-Need-Measure-2025-_csv.csv)
        
        if [ -n "$LAT" ] && [ -n "$LNG" ]; then
            echo "{c:\"$LSOA\",n:\"$NAME\",lat:$LAT,lng:$LNG,m:\"$MISSION\"},"
        fi
    done
    
    echo "],"
done

echo "};"
