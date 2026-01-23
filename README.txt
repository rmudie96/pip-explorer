PRIDE IN PLACE DATA EXPLORER

THE PROBLEM:
The map loads a 200MB boundary file. This is too big for browsers to handle.
We need ALL 191 LSOAs to display with boundaries.

THE SOLUTION:
Create a smaller boundary file with ONLY our 191 LSOAs (should be ~2-5MB instead of 200MB)

STATUS:
- App works perfectly with embedded data ✓
- Need to extract just our boundaries from the big file
- Once we have the small file, we change one line in index.html to load it
- Then it will work online

WHAT NEEDS TO HAPPEN:
1. Extract 191 LSOA boundaries from the big file → create small file
2. Upload small file with index.html
3. Done - will work online