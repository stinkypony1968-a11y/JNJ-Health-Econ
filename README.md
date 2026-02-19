# J&J MedTech Neurovascular — Territory Map Generator

Generate single-file HTML territory maps from Excel hospital data. Each output is a self-contained, interactive map with Leaflet.js — ready for Netlify deployment with zero build step.

![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue)

## What It Does

Takes an `.xlsx` file with hospital data and a small JSON config, and outputs a complete territory map HTML file with:

- **Interactive Leaflet map** — CartoDB light tiles, custom div markers sized by NTS and colored by stroke certification (CSC=red, PSC=amber, TSC/TCC=blue, None=gray)
- **Left sidebar** — Hospital cards ranked by NTS, searchable, filterable by cert type, sortable by multiple fields
- **Account detail drawer** — Full hospital profile, TPS score gauge, stroke prevalence bars, auto-generated business plan, travel logistics
- **Cadence calendar** — Mar–Aug 2026, per-hospital configurable cadence with department-colored dots (Physician Lab, Stroke Coord, Education Coord, Buying/VAC)
- **Bottom analytics bar** — Territory NTS totals, projected growth, TPS distribution, tier breakdown
- **Health economic enrichment** — LVO, cSDH/MMA, and hemorrhagic stroke prevalence rates auto-calculated from catchment population
- **TPS auto-scoring** — 0–100 Territory Priority Score from 6 weighted factors (stroke volume, beds, certification, fellowships, sales momentum, revenue rank)

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Generate with full config
python generate_territory_map.py --excel hospitals.xlsx --config config.json --output territory-map.html

# Generate with minimal CLI args
python generate_territory_map.py hospitals.xlsx --rep "Andrew Payne" --territory "Florida / Georgia" --output flga-map.html
```

## Excel Input Format

Your `.xlsx` file needs one row per hospital with these columns:

### Required

| Column | Example |
|--------|---------|
| Hospital Name | Baptist Medical Center Jacksonville |
| Short Name | Baptist Jacksonville |
| Address | 800 Prudential Dr, Jacksonville, FL 32207 |
| City | Jacksonville |
| State | FL |
| County | Duval County, FL |
| Latitude | 30.3187 |
| Longitude | -81.6568 |
| Phone | (904) 202-2000 |
| Licensed Beds | 972 |
| Affiliation | Baptist Health |
| Buy Group/GPO | Vizient |
| Stroke Certification | CSC / PSC / TSC / TCC / None |
| Strokes Per Year | 900 |
| NTS 2025 (Full Year) | 26327 |
| NTS 2026 (YTD) | 46025 |

### Optional

| Column | Example |
|--------|---------|
| IR Phone | (904) 202-2100 |
| Cert Body | Joint Commission / DNV |
| Neurosurgery Fellowship | TRUE / FALSE |
| Neuro IR Fellowship | TRUE / FALSE |
| Catchment Population | 1750000 |
| Median Age | 37.5 |
| Life Expectancy | 77.4 |
| Comment | 175% YoY growth |

See `sample_hospitals.xlsx` for a working example with 10 FL/GA hospitals.

## Config JSON

```json
{
  "rep_name": "Andrew Payne",
  "rep_role": "CAS",
  "rep_base_city": "Jacksonville, FL",
  "rep_base_lat": 30.3322,
  "rep_base_lng": -81.6557,
  "territory_name": "Florida / Georgia",
  "map_center_lat": 29.80,
  "map_center_lng": -81.80,
  "map_zoom": 7,
  "team_members": [
    { "name": "Mark", "role": "TM", "city": "Orlando, FL", "lat": 28.5383, "lng": -81.3792 }
  ]
}
```

## Auto-Calculations

### Health Economics (from catchment population)
- **LVO Rate**: population × 24 / 100,000 *(Rai et al., Stroke 2020)*
- **cSDH/MMA Rate**: population × 17.3 / 100,000 *(Rauhala et al., JAMA Neurology 2020)*
- **Hemorrhagic Rate**: population × 12 / 100,000 *(GBD 2019 / AHA 2024)*

### TPS Scoring (0–100)
| Factor | Max Points |
|--------|-----------|
| Stroke volume (vs territory max) | 25 |
| Certification (CSC=20, PSC=12, TSC/TCC=8, None=3) | 20 |
| Bed count (vs territory max) | 15 |
| Fellowship programs (NSG=7.5, NIR=7.5) | 15 |
| Sales momentum (growing=15, flat=8, declining=3) | 15 |
| Revenue quartile rank (Q1=10, Q2=7, Q3=4, Q4=2) | 10 |

Segments: **HIGH** (70+) → Invest & Grow, **MID** (50–69) → Develop & Build, **LOW** (<50) → Monitor & Nurture

### Tier Assignment
Hospitals ranked by NTS 2025 descending. Top third = Tier 1, middle = Tier 2, bottom = Tier 3.

### Travel Time
Haversine distance × 1.3 road factor ÷ 55 mph, calculated from the closest team member.

## Project Structure

```
├── generate_territory_map.py   # Main generator script
├── territory_template.html     # Jinja2 HTML template (all CSS/JS inline)
├── sample_config.json          # Example config file
├── sample_hospitals.xlsx       # 10 sample FL/GA hospitals
├── requirements.txt            # Python dependencies
└── README.md
```

## Deployment

The output HTML is a single file with no build step. Drop it on Netlify, Vercel, or any static host:

```bash
# Generate
python generate_territory_map.py --excel hospitals.xlsx --config config.json --output dist/index.html

# Deploy to Netlify
netlify deploy --prod --dir=dist
```

## Branding

Uses J&J MedTech design tokens:
- Navy: `#00205B`
- Red: `#EB1700`
- Blue: `#0077C8`
- Font: Inter (Google Fonts CDN)

## License

Internal use — J&J MedTech Neurovascular territory management.
