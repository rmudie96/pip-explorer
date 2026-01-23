# Pride in Place Data Explorer

An interactive neighbourhood intelligence platform for the Pride in Place programme, providing comprehensive data visualisation and analysis for community development.

## Features

- **Interactive Map Visualization**: Explore Lower-layer Super Output Areas (LSOAs) with color-coded mission classifications
- **Hyper-Local Need Measure (HLNM)**: Five key indicators measuring growth, energy, crime, opportunity, and health needs
- **Community Needs Index (CNI)**: Measures of civic infrastructure, community connectedness, and civic engagement
- **Deprivation Metrics**: English Indices of Multiple Deprivation (IMD 2025) converted to accessible percentile rankings
- **Mission Framework**: Critical, Priority, and Support classifications for targeted intervention

## Quick Start

### View Online
Visit the live site: [Coming soon - will be available after GitHub Pages deployment]

### Run Locally

1. Clone this repository
2. Start a local web server in the project directory:

   **Using Python:**
   ```bash
   python -m http.server 8000
   ```

   **Using Node.js:**
   ```bash
   npx http-server
   ```

3. Open http://localhost:8000 in your browser

**Important:** The map boundaries require a web server to load properly due to browser security restrictions. Simply opening `index.html` directly will only show location markers, not boundary polygons.

## Data Sources

- **Census 2021**: ONS Census data via NOMIS
- **IMD 2025**: English Indices of Multiple Deprivation
- **HLNM 2024**: OCSI Hyper-Local Need Measure
- **CNI 2023**: OCSI Community Needs Index
- **Boundaries**: ONS Geography Portal (LSOA and MSOA boundaries)

## Coverage

Currently covers **40 Pride in Place neighbourhoods** across England, including areas in:
- Birmingham
- Liverpool
- Manchester
- Leeds
- Kingston upon Hull
- Leicester
- North Northamptonshire
- Bristol
- And more...

## Technology

Built with vanilla JavaScript, Leaflet.js for mapping, and modern CSS. No build process required - just open and run.

## Project Structure

- `index.html` - Main application (fully self-contained)
- `*.geojson` - Boundary data files for map visualization
- `lsoa_embedded_data_temp.js` - LSOA classification data
- `data.json` - Source data for all neighbourhoods
- Various `.ps1` and `.py` scripts for data processing

## Development

See `DEVELOPMENT_GUIDE.md` for details on updating data sources and maintaining the platform.

## License

Created for Local Trust's Pride in Place programme.

## Contact

For questions or feedback about this platform, please contact the Local Trust team.

---

*Built with ❤️ for community empowerment and data-driven local development*
