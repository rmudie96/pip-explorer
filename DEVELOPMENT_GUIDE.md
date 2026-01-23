# Pride in Place Data Explorer - Development Guide

## Project Overview

The Pride in Place Data Explorer is an interactive web-based dashboard for exploring socioeconomic indicators for Pride in Place Programme neighbourhoods in England. Built for Local Trust/ICON.

---

## Session Log

### Session 1: Initial Build (January 2026)

**Objective:** Create the foundational data explorer with Census 2021 and IMD 2025 data.

**Steps Completed:**
1. Created `data_gatherer.py` - Python script to download and process data from:
   - OCSI (Pride in Place MSOA list)
   - ONS Census 2021 via NOMIS (economic activity, health, education, tenure, etc.)
   - MHCLG English Indices of Deprivation 2025
2. Built `index.html` - Self-contained interactive dashboard with:
   - Dropdown selector for 40 Pride in Place neighbourhoods
   - Search/filter functionality
   - Colour-coded percentile bars (quintile system)
   - Categories: Employment, Health, Education, Housing, Deprivation, IMD Scores
   - Relative strengths and key challenges summary

**Data Sources Used:**
- Pride in Place MSOA list from OCSI
- Census 2021 bulk downloads from NOMIS
- IMD 2025 from GOV.UK

---

### Session 2: Feature Expansion (21 January 2026)

**Objective:** Expand the explorer with additional datasets and features.

**Tasks:**
1. **Add HLNM (Hyper-Local Need Measure) datasets** - All indicators from HLNM_MSOA.xlsx
   - Higher number = worse for all indicators except crime (where lower = worse)

2. **LSOA Breakdown Feature** - Using Hyper-Local-Need-Measure-2025.xlsx
   - Show how many LSOAs within each MSOA are "Mission Critical", "Mission Priority", "Mission Support"
   - Clickable drill-down to see individual LSOA data

3. **Community Needs Index** - Add 4 new datapoints at top of dashboard:
   - Community Needs Rank (overall index)
   - Active and Engaged Community rank (subdomain)
   - Civic Assets rank (subdomain)
   - Connectedness rank (subdomain)

4. **Local Authority Search & Comparison**
   - Search by Local Authority
   - View multiple places simultaneously
   - Compare areas side-by-side

5. **Benchmarking Feature**
   - Compare each area against LA average, regional average, national average

6. **Citations**
   - Add data source citations under each data piece

7. **Download Functionality**
   - Allow users to download data for selected areas

8. **Branding Updates**
   - Integrate ICON logo
   - Apply ICON colour theme:
     - Primary Dark: #002B5C
     - Primary Teal: #009B9F
     - Light Blue: #B0E0E6
     - Pale Blue: #D9E8F2
     - Purple/Blue: #6E7BA2
     - Dark Grey: #4A4A4A
   - Apply Abadi fonts (Bold for headers, Extra Light for body)

---

## Technical Architecture

### File Structure
```
PiP Explorer/
├── index.html              # Main dashboard (self-contained)
├── data.json               # Generated data file
├── data_gatherer.py        # Python data processing script
├── DEVELOPMENT_GUIDE.md    # This file
├── data/                   # Cached downloaded data
├── HLNM_MSOA.xlsx         # Hyper-Local Need Measure at MSOA level
├── Hyper-Local-Need-Measure-2025.xlsx  # LSOA-level data
├── MSOA_Community Needs Index 2023_*.xlsx  # CNI files
├── ICON-Logo-Final_optimised.png  # Company logo
└── ICON Colour Theme.xml   # Brand colours
```

### Data Processing Pipeline
1. Download source data from official APIs/websites
2. Process and aggregate to MSOA level
3. Calculate percentile rankings against all England MSOAs
4. Export to JSON format
5. Embed in HTML dashboard

---

## Colour Scheme (ICON Brand)

| Name | Hex Code | Usage |
|------|----------|-------|
| Primary Dark | #002B5C | Headers, primary buttons |
| Primary Teal | #009B9F | Accents, highlights |
| Light Blue | #B0E0E6 | Backgrounds, cards |
| Pale Blue | #D9E8F2 | Light backgrounds |
| Purple/Blue | #6E7BA2 | Links, secondary elements |
| Dark Grey | #4A4A4A | Body text |
| White | #FFFFFF | Backgrounds |
| Light Grey | #E7E6E6 | Borders, dividers |

## Typography

- **Headers/Bold Text:** Abadi (Bold)
- **Body/Light Text:** Abadi Extra Light

---

## Data Sources & Citations

| Dataset | Source | URL |
|---------|--------|-----|
| Pride in Place MSOA List | OCSI | https://ocsi.uk |
| Census 2021 | ONS via NOMIS | https://www.nomisweb.co.uk |
| IMD 2025 | MHCLG | https://www.gov.uk/government/statistics/english-indices-of-deprivation-2025 |
| Hyper-Local Need Measure | OCSI | https://ocsi.uk |
| Community Needs Index | OCSI | https://ocsi.uk |

---

## Future Enhancements (Planned)

- [ ] Add map visualisation
- [ ] Time series comparison (when data available)
- [ ] Export to PDF reports
- [ ] API endpoint for programmatic access

---

*Last Updated: 21 January 2026*
