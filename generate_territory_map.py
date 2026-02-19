#!/usr/bin/env python3
"""
Territory Map Generator for J&J MedTech Neurovascular
=====================================================
Generates a single-file HTML territory map from an Excel file of hospital data.

Usage:
    python generate_territory_map.py --excel hospitals.xlsx --config config.json --output territory-map.html
    python generate_territory_map.py hospitals.xlsx --rep "Andrew Payne" --territory "Florida / Georgia" --output flga-map.html
"""

import argparse
import json
import math
import os
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
from jinja2 import Template

# ─── Constants ───────────────────────────────────────────────────────────────

REQUIRED_COLUMNS = [
    "Hospital Name", "Short Name", "Address", "City", "State", "County",
    "Latitude", "Longitude", "Phone", "Licensed Beds", "Affiliation",
    "Buy Group/GPO", "Stroke Certification", "Strokes Per Year",
    "NTS 2025 (Full Year)", "NTS 2026 (YTD)"
]

OPTIONAL_COLUMNS = [
    "IR Phone", "Cert Body", "Neurosurgery Fellowship", "Neuro IR Fellowship",
    "Catchment Population", "Median Age", "Life Expectancy", "Comment"
]

CERT_SCORES = {"CSC": 20, "PSC": 12, "TSC": 8, "TCC": 8, "None": 3, "": 3}
CERT_COLORS = {"CSC": "#EB1700", "PSC": "#F59E0B", "TSC": "#0077C8", "TCC": "#0077C8", "None": "#6B7280", "": "#6B7280"}
CERT_LABELS = {"CSC": "Comprehensive Stroke Center", "PSC": "Primary Stroke Center",
               "TSC": "Thrombectomy-Capable Stroke Center", "TCC": "Thrombectomy-Capable Center",
               "None": "No Certification", "": "No Certification"}

SEGMENT_THRESHOLDS = {"HIGH": 70, "MID": 50}
SEGMENT_LABELS = {"HIGH": "Invest & Grow", "MID": "Develop & Build", "LOW": "Monitor & Nurture"}

# ─── Data Loading ────────────────────────────────────────────────────────────

def load_excel(filepath):
    """Load hospital data from Excel file."""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active

    # Read headers from first row
    headers = []
    for cell in ws[1]:
        headers.append(str(cell.value).strip() if cell.value else "")

    # Validate required columns
    missing = [c for c in REQUIRED_COLUMNS if c not in headers]
    if missing:
        print(f"ERROR: Missing required columns: {', '.join(missing)}")
        print(f"Found columns: {', '.join(headers)}")
        sys.exit(1)

    # Read data rows
    hospitals = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:  # Skip empty rows
            continue
        record = {}
        for i, h in enumerate(headers):
            if i < len(row):
                record[h] = row[i]
            else:
                record[h] = None
        hospitals.append(record)

    wb.close()
    return hospitals


def load_config(config_path=None, args=None):
    """Load configuration from JSON file or command-line args."""
    config = {
        "rep_name": "Territory Rep",
        "rep_role": "CAS",
        "rep_base_city": "",
        "rep_base_lat": 0,
        "rep_base_lng": 0,
        "territory_name": "Territory",
        "map_center_lat": 33.0,
        "map_center_lng": -83.0,
        "map_zoom": 7,
        "team_members": []
    }

    if config_path and os.path.exists(config_path):
        with open(config_path, "r") as f:
            file_config = json.load(f)
        config.update(file_config)

    # CLI overrides
    if args:
        if args.rep:
            config["rep_name"] = args.rep
        if args.territory:
            config["territory_name"] = args.territory

    return config


# ─── Data Enrichment ─────────────────────────────────────────────────────────

def safe_float(val, default=0.0):
    """Safely convert to float."""
    if val is None:
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def safe_int(val, default=0):
    """Safely convert to int."""
    if val is None:
        return default
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return default


def safe_bool(val, default=False):
    """Safely convert to bool."""
    if val is None:
        return default
    if isinstance(val, bool):
        return val
    if isinstance(val, str):
        return val.lower() in ("true", "yes", "1", "y")
    return bool(val)


def enrich_hospital(h):
    """Add computed fields to a hospital record."""
    # Normalize fields
    h["lat"] = safe_float(h.get("Latitude"))
    h["lng"] = safe_float(h.get("Longitude"))
    h["beds"] = safe_int(h.get("Licensed Beds"))
    h["strokes_yr"] = safe_int(h.get("Strokes Per Year"))
    h["nts_2025"] = safe_int(h.get("NTS 2025 (Full Year)"))
    h["nts_2026_ytd"] = safe_int(h.get("NTS 2026 (YTD)"))
    h["catchment"] = safe_int(h.get("Catchment Population"))
    h["median_age"] = safe_float(h.get("Median Age"))
    h["life_expectancy"] = safe_float(h.get("Life Expectancy"))
    h["nsg_fellowship"] = safe_bool(h.get("Neurosurgery Fellowship"))
    h["nir_fellowship"] = safe_bool(h.get("Neuro IR Fellowship"))
    h["cert"] = str(h.get("Stroke Certification", "None")).strip()
    if h["cert"] not in CERT_SCORES:
        h["cert"] = "None"
    h["cert_body"] = str(h.get("Cert Body", "")) if h.get("Cert Body") else ""
    h["comment"] = str(h.get("Comment", "")) if h.get("Comment") else ""
    h["ir_phone"] = str(h.get("IR Phone", "")) if h.get("IR Phone") else ""
    h["name"] = str(h.get("Hospital Name", ""))
    h["short_name"] = str(h.get("Short Name", ""))
    h["address"] = str(h.get("Address", ""))
    h["city"] = str(h.get("City", ""))
    h["state"] = str(h.get("State", ""))
    h["county"] = str(h.get("County", ""))
    h["phone"] = str(h.get("Phone", ""))
    h["affiliation"] = str(h.get("Affiliation", ""))
    h["gpo"] = str(h.get("Buy Group/GPO", ""))

    # Health economic enrichment
    if h["catchment"] > 0:
        h["lvo_rate"] = round(h["catchment"] * 24 / 100000)
        h["csdh_mma_rate"] = round(h["catchment"] * 17.3 / 100000)
        h["hemorrhagic_rate"] = round(h["catchment"] * 12 / 100000)
    else:
        h["lvo_rate"] = 0
        h["csdh_mma_rate"] = 0
        h["hemorrhagic_rate"] = 0

    # Annualized 2026 (YTD as of ~mid-Feb, so ~1.5 months = 1.5/12 factor)
    # Use a dynamic fraction based on current month
    now = datetime.now()
    month_fraction = (now.month - 1 + now.day / 30) / 12
    if month_fraction > 0:
        h["nts_2026_annualized"] = round(h["nts_2026_ytd"] / month_fraction)
    else:
        h["nts_2026_annualized"] = h["nts_2026_ytd"]

    # YoY growth
    if h["nts_2025"] > 0:
        h["yoy_growth"] = round((h["nts_2026_annualized"] - h["nts_2025"]) / h["nts_2025"] * 100, 1)
    else:
        h["yoy_growth"] = 0

    # Cert color and label
    h["cert_color"] = CERT_COLORS.get(h["cert"], "#6B7280")
    h["cert_label"] = CERT_LABELS.get(h["cert"], "No Certification")

    return h


def calculate_tps(hospitals):
    """Calculate TPS scores for all hospitals."""
    if not hospitals:
        return hospitals

    max_strokes = max(h["strokes_yr"] for h in hospitals) or 1
    max_beds = max(h["beds"] for h in hospitals) or 1

    # Revenue quartiles
    nts_values = sorted([h["nts_2025"] for h in hospitals], reverse=True)
    q1_threshold = nts_values[len(nts_values) // 4] if len(nts_values) >= 4 else nts_values[0]
    q2_threshold = nts_values[len(nts_values) // 2] if len(nts_values) >= 2 else nts_values[0]
    q3_threshold = nts_values[3 * len(nts_values) // 4] if len(nts_values) >= 4 else 0

    for h in hospitals:
        score = 0

        # Stroke volume (0-25)
        score += (h["strokes_yr"] / max_strokes) * 25

        # Bed count (0-15)
        score += (h["beds"] / max_beds) * 15

        # Stroke certification (0-20)
        score += CERT_SCORES.get(h["cert"], 3)

        # Fellowship programs (0-15)
        if h["nsg_fellowship"]:
            score += 7.5
        if h["nir_fellowship"]:
            score += 7.5

        # Sales momentum (0-15)
        if h["nts_2026_annualized"] > h["nts_2025"] * 1.05:
            score += 15
        elif h["nts_2026_annualized"] >= h["nts_2025"] * 0.95:
            score += 8
        else:
            score += 3

        # Revenue rank (0-10)
        if h["nts_2025"] >= q1_threshold:
            score += 10
        elif h["nts_2025"] >= q2_threshold:
            score += 7
        elif h["nts_2025"] >= q3_threshold:
            score += 4
        else:
            score += 2

        h["tps_score"] = round(min(score, 100))

        # Segment
        if h["tps_score"] >= SEGMENT_THRESHOLDS["HIGH"]:
            h["tps_segment"] = "HIGH"
        elif h["tps_score"] >= SEGMENT_THRESHOLDS["MID"]:
            h["tps_segment"] = "MID"
        else:
            h["tps_segment"] = "LOW"

        h["tps_strategy"] = SEGMENT_LABELS[h["tps_segment"]]

    return hospitals


def assign_tiers(hospitals):
    """Assign tiers based on NTS 2025 ranking."""
    sorted_h = sorted(hospitals, key=lambda x: x["nts_2025"], reverse=True)
    n = len(sorted_h)
    t1 = max(1, n // 3)
    t2 = max(2, 2 * n // 3)

    for i, h in enumerate(sorted_h):
        if i < t1:
            h["tier"] = 1
        elif i < t2:
            h["tier"] = 2
        else:
            h["tier"] = 3

    return hospitals


def calculate_travel(hospitals, config):
    """Calculate travel time from closest team member using Haversine."""
    team_locations = [{"lat": config["rep_base_lat"], "lng": config["rep_base_lng"],
                       "name": config["rep_name"], "role": config["rep_role"]}]
    for tm in config.get("team_members", []):
        team_locations.append({"lat": tm["lat"], "lng": tm["lng"],
                               "name": tm["name"], "role": tm["role"]})

    for h in hospitals:
        min_dist = float("inf")
        closest = team_locations[0]["name"]
        for tl in team_locations:
            d = haversine(h["lat"], h["lng"], tl["lat"], tl["lng"])
            if d < min_dist:
                min_dist = d
                closest = tl["name"]

        road_miles = min_dist * 1.3  # road factor
        h["travel_miles"] = round(road_miles, 1)
        h["travel_time_min"] = round(road_miles / 55 * 60)  # 55 mph
        h["closest_team_member"] = closest

    return hospitals


def haversine(lat1, lon1, lat2, lon2):
    """Calculate distance in miles between two coordinates."""
    R = 3959  # Earth's radius in miles
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat / 2) ** 2 +
         math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) *
         math.sin(dlon / 2) ** 2)
    return R * 2 * math.asin(math.sqrt(a))


# ─── Template Rendering ─────────────────────────────────────────────────────

def render_template(hospitals, config, template_path):
    """Render the HTML template with hospital data and config."""
    with open(template_path, "r") as f:
        template_str = f.read()

    template = Template(template_str)

    # Sort hospitals by NTS 2025 descending for default display
    hospitals_sorted = sorted(hospitals, key=lambda x: x["nts_2025"], reverse=True)

    # Compute territory-level stats
    total_nts_2025 = sum(h["nts_2025"] for h in hospitals)
    total_nts_2026_ytd = sum(h["nts_2026_ytd"] for h in hospitals)
    total_nts_2026_ann = sum(h["nts_2026_annualized"] for h in hospitals)
    territory_yoy = round((total_nts_2026_ann - total_nts_2025) / total_nts_2025 * 100, 1) if total_nts_2025 > 0 else 0

    tier_counts = {1: 0, 2: 0, 3: 0}
    tps_dist = {"HIGH": 0, "MID": 0, "LOW": 0}
    cert_counts = {"CSC": 0, "PSC": 0, "TSC": 0, "TCC": 0, "None": 0}
    for h in hospitals:
        tier_counts[h["tier"]] = tier_counts.get(h["tier"], 0) + 1
        tps_dist[h["tps_segment"]] = tps_dist.get(h["tps_segment"], 0) + 1
        cert_key = h["cert"] if h["cert"] in cert_counts else "None"
        cert_counts[cert_key] = cert_counts.get(cert_key, 0) + 1

    avg_tps = round(sum(h["tps_score"] for h in hospitals) / len(hospitals)) if hospitals else 0

    generated_date = datetime.now().strftime("%B %d, %Y")

    html = template.render(
        hospitals=hospitals_sorted,
        hospitals_json=json.dumps([serialize_hospital(h) for h in hospitals_sorted]),
        config=config,
        total_nts_2025=total_nts_2025,
        total_nts_2026_ytd=total_nts_2026_ytd,
        total_nts_2026_ann=total_nts_2026_ann,
        territory_yoy=territory_yoy,
        tier_counts=tier_counts,
        tps_dist=tps_dist,
        cert_counts=cert_counts,
        avg_tps=avg_tps,
        hospital_count=len(hospitals),
        generated_date=generated_date,
    )
    return html


def serialize_hospital(h):
    """Convert hospital to JSON-safe dict."""
    keys = [
        "name", "short_name", "address", "city", "state", "county",
        "lat", "lng", "phone", "ir_phone", "beds", "affiliation", "gpo",
        "cert", "cert_body", "cert_color", "cert_label",
        "nsg_fellowship", "nir_fellowship", "strokes_yr",
        "catchment", "median_age", "life_expectancy",
        "lvo_rate", "csdh_mma_rate", "hemorrhagic_rate",
        "nts_2025", "nts_2026_ytd", "nts_2026_annualized", "yoy_growth",
        "tps_score", "tps_segment", "tps_strategy",
        "tier", "travel_miles", "travel_time_min", "closest_team_member",
        "comment"
    ]
    return {k: h.get(k) for k in keys}


# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Generate J&J MedTech Neurovascular Territory Map from Excel data"
    )
    parser.add_argument("excel", nargs="?", help="Path to hospital Excel file")
    parser.add_argument("--excel", dest="excel_flag", help="Path to hospital Excel file")
    parser.add_argument("--config", help="Path to config JSON file")
    parser.add_argument("--output", "-o", default="territory-map.html", help="Output HTML file path")
    parser.add_argument("--rep", help="Rep name (overrides config)")
    parser.add_argument("--territory", help="Territory name (overrides config)")
    parser.add_argument("--template", default=None, help="Path to HTML template (default: territory_template.html in same dir)")

    args = parser.parse_args()

    # Resolve Excel path
    excel_path = args.excel or args.excel_flag
    if not excel_path:
        parser.error("Excel file is required. Use: generate_territory_map.py hospitals.xlsx")

    if not os.path.exists(excel_path):
        print(f"ERROR: Excel file not found: {excel_path}")
        sys.exit(1)

    # Resolve template path
    script_dir = Path(__file__).parent
    template_path = args.template or str(script_dir / "territory_template.html")
    if not os.path.exists(template_path):
        print(f"ERROR: Template file not found: {template_path}")
        sys.exit(1)

    # Load data
    print(f"Loading hospitals from {excel_path}...")
    hospitals = load_excel(excel_path)
    print(f"  Loaded {len(hospitals)} hospitals")

    # Load config
    config = load_config(args.config, args)
    print(f"  Territory: {config['territory_name']}")
    print(f"  Rep: {config['rep_name']} ({config['rep_role']})")

    # Enrich data
    print("Enriching hospital data...")
    hospitals = [enrich_hospital(h) for h in hospitals]

    # Calculate TPS scores
    print("Calculating TPS scores...")
    hospitals = calculate_tps(hospitals)

    # Assign tiers
    print("Assigning tiers...")
    hospitals = assign_tiers(hospitals)

    # Calculate travel times
    print("Calculating travel times...")
    hospitals = calculate_travel(hospitals, config)

    # Auto-detect map center if not set
    if config["map_center_lat"] == 0 and config["map_center_lng"] == 0:
        config["map_center_lat"] = sum(h["lat"] for h in hospitals) / len(hospitals)
        config["map_center_lng"] = sum(h["lng"] for h in hospitals) / len(hospitals)

    # Render template
    print(f"Rendering template from {template_path}...")
    html = render_template(hospitals, config, template_path)

    # Write output
    with open(args.output, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\nOutput written to {args.output}")
    print(f"  File size: {os.path.getsize(args.output) / 1024:.1f} KB")

    # Summary
    print("\n── Territory Summary ──────────────────────────")
    print(f"  Hospitals:     {len(hospitals)}")
    tier_1 = sum(1 for h in hospitals if h['tier'] == 1)
    tier_2 = sum(1 for h in hospitals if h['tier'] == 2)
    tier_3 = sum(1 for h in hospitals if h['tier'] == 3)
    print(f"  Tier 1 / 2 / 3: {tier_1} / {tier_2} / {tier_3}")
    high = sum(1 for h in hospitals if h['tps_segment'] == 'HIGH')
    mid = sum(1 for h in hospitals if h['tps_segment'] == 'MID')
    low = sum(1 for h in hospitals if h['tps_segment'] == 'LOW')
    print(f"  TPS HIGH/MID/LOW: {high} / {mid} / {low}")
    total = sum(h['nts_2025'] for h in hospitals)
    print(f"  Total NTS 2025: ${total:,}")
    total_26 = sum(h['nts_2026_annualized'] for h in hospitals)
    print(f"  NTS 2026 (ann.): ${total_26:,}")
    print(f"  Map center:    ({config['map_center_lat']:.4f}, {config['map_center_lng']:.4f})")
    print(f"  Map zoom:      {config['map_zoom']}")
    print("───────────────────────────────────────────────")


if __name__ == "__main__":
    main()
