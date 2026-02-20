# J&J MedTech Neurovascular — Territory Map Generator

## Health Economics & Clinical Infrastructure Edition

Generate single-file HTML territory maps from Excel hospital data. Each output is a self-contained, interactive map with Leaflet.js — focused on stroke epidemiology, clinical infrastructure, and geographic access. Ready for Netlify deployment with zero build step.

![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue)

## What It Does

Takes an `.xlsx` file with hospital data and a small JSON config, and outputs a complete territory map HTML file with:

- **Interactive Leaflet map** — CartoDB light tiles, custom div markers sized by stroke volume and colored by stroke certification (CSC=red, PSC=amber, TSC/TCC=blue, None=gray)
- **Left sidebar** — Hospital cards ranked by infrastructure score, searchable, filterable by cert type, sortable by multiple fields
- **Account detail drawer** — Full hospital profile, infrastructure score gauge, infrastructure checklist (✓/✗), stroke epidemiology panel, A-Fib data, geographic access & travel
- **Cadence calendar** — Mar–Aug 2026, per-hospital configurable cadence with department-colored dots
- **Bottom analytics bar** — Territory beds, stroke volume, LVO, MT eligible, cert breakdown, capability counts, tier distribution
- **Stroke epidemiology** — Ischemic, LVO, MT eligible, hemorrhagic subtypes, AVM/aneurysm volumes from effective population
- **Infrastructure scoring** — 0–100 score from certification, capabilities, fellowships, and stroke volume

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

### Optional — Clinical Infrastructure

| Column | Example |
|--------|---------|
| CAH Status | TRUE / FALSE |
| Telestroke Capable | TRUE / FALSE |
| 24/7 Thrombectomy | TRUE / FALSE |
| tPA Available | TRUE / FALSE |
| Neuro ICU | TRUE / FALSE |
| CT Scanner | TRUE / FALSE |
| Spoke Hospital | TRUE / FALSE |

### Optional — Demographics & Epidemiology

| Column | Example |
|--------|---------|
| Catchment Population | 1750000 |
| Population 65+ | 310000 |
| Median Age | 37.5 |
| Life Expectancy | 77.4 |
| Neurosurgery Fellowship | TRUE / FALSE |
| Neuro IR Fellowship | TRUE / FALSE |
| IR Phone | (904) 202-2100 |
| Cert Body | Joint Commission / DNV |

### Optional — A-Fib & Geographic Access

| Column | Example |
|--------|---------|
| A-Fib Prevalence | 3.2 |
| A-Fib Absolute Count | 56000 |
| MMAE Score | 72 |
| Clinical Benefit Radius (km) | 45 |
| Medevac Available | TRUE / FALSE |
| Road Access | TRUE / FALSE |
| Comment | Flagship CSC |

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

### Stroke Epidemiology (from catchment population)

Effective population is calculated as: `pop_under_65 + (pop_65_plus × 2.5)`

| Metric | Rate | Source |
|--------|------|--------|
| Ischemic stroke | 216/100k effective pop | Rai et al., Stroke 2020 |
| LVO | 21% of ischemic | Literature consensus |
| MT eligibility | 70% of LVO | DEFUSE/DAWN trials |
| AVM/Aneurysm | 12/100k effective pop | GBD 2019 |
| Hemorrhagic | 12/100k effective pop | AHA 2024 |
| SAH | 5% of hemorrhagic | — |
| ICH | 80% of hemorrhagic | — |

### Infrastructure Score (0–100)

| Factor | Max Points |
|--------|-----------|
| Certification (CSC=25, PSC=20, TSC/TCC=15, None=5) | 25 |
| 24/7 Thrombectomy | 15 |
| Neuro ICU | 12 |
| Critical Access Hospital | 10 |
| CT Scanner | 8 |
| Telestroke | 8 |
| Neuro IR Fellowship | 8 |
| Neurosurgery Fellowship | 7 |
| tPA Available | 5 |
| Stroke volume bonus (vs territory max) | 10 |

### Clinical Priority Tiers

Hospitals ranked by infrastructure score + stroke volume descending. Top third = Tier 1 (High Complexity Hub), middle = Tier 2 (Regional Center), bottom = Tier 3 (Basic Capability).

### Geographic Access

Computed from travel time: Local (<30 min), Short (30–120), Long (120–240), Flight Required (>240).

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
