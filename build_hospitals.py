#!/usr/bin/env python3
"""Build sample_hospitals.xlsx with all 21 FL/GA territory hospitals."""
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Hospitals"

headers = [
    "Hospital Name", "Short Name", "Address", "City", "State", "County",
    "Latitude", "Longitude", "Phone", "Licensed Beds", "Affiliation",
    "Buy Group/GPO", "Stroke Certification", "Strokes Per Year",
    "IR Phone", "Cert Body", "Neurosurgery Fellowship", "Neuro IR Fellowship",
    "Catchment Population", "Population 65+", "Median Age", "Life Expectancy",
    "CAH Status", "Telestroke Capable", "24/7 Thrombectomy",
    "tPA Available", "Neuro ICU", "CT Scanner", "Spoke Hospital",
    "A-Fib Prevalence", "A-Fib Absolute Count", "MMAE Score",
    "Clinical Benefit Radius (km)",
    "Medevac Available", "Road Access", "Comment"
]

header_fill = PatternFill(start_color="00205B", end_color="00205B", fill_type="solid")
header_font = Font(bold=True, size=10, color="FFFFFF")

for col, h in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", wrap_text=True)

# ── Hospital Data ──────────────────────────────────────────────────────
# Each row: [name, short, address, city, state, county, lat, lng, phone, beds,
#            affiliation, gpo, cert, strokes_yr,
#            ir_phone, cert_body, nsg_fellowship, nir_fellowship,
#            pop, pop65, median_age, life_exp,
#            cah, telestroke, thrombectomy_24_7, tpa, neuro_icu, ct, spoke,
#            afib_prev, afib_count, mmae, benefit_radius,
#            medevac, road_access, comment]

hospitals = [
    # 1 - 00004823 SAVANNAH HEALTH SERVICES LLC (Memorial Health Savannah)
    ["Savannah Health Services LLC (Memorial Health)", "Memorial Savannah",
     "4700 Waters Ave, Savannah, GA 31404", "Savannah", "GA", "Chatham County, GA",
     32.0312, -81.0874, "(912) 350-8000", 612, "HCA Healthcare", "HealthTrust",
     "CSC", 580,
     "(912) 350-8200", "Joint Commission", True, False,
     680000, 108800, 35.7, 75.2,
     False, True, True, True, True, True, False,
     2.7, 18360, 62, 35,
     True, True, "Regional CSC hub for southeast Georgia"],

    # 2 - 01544090 ADVENT HEALTH CSC NON STOCK (AdventHealth — likely Celebration/Altamonte)
    ["AdventHealth CSC", "AdventHealth CSC",
     "601 E Rollins St, Orlando, FL 32803", "Orlando", "FL", "Orange County, FL",
     28.5597, -81.3628, "(407) 303-5600", 547, "AdventHealth", "Premier",
     "CSC", 620,
     "(407) 303-5700", "Joint Commission", True, True,
     1200000, 192000, 36.4, 78.8,
     False, True, True, True, True, True, False,
     3.0, 36000, 70, 40,
     True, True, "AdventHealth system CSC — strong neuro IR program"],

    # 3 - 00002905 BAPTIST MEDICAL CENTER (Jacksonville)
    ["Baptist Medical Center Jacksonville", "Baptist Jax",
     "800 Prudential Dr, Jacksonville, FL 32207", "Jacksonville", "FL", "Duval County, FL",
     30.3187, -81.6568, "(904) 202-2000", 972, "Baptist Health", "Vizient",
     "CSC", 900,
     "(904) 202-2100", "Joint Commission", True, True,
     1750000, 310000, 37.5, 77.4,
     False, True, True, True, True, True, False,
     3.2, 56000, 72, 45,
     True, True, "Flagship CSC — largest stroke program in territory"],

    # 4 - 01110024 SHANDS TEACHING HOSPITAL (UF Health Shands Gainesville)
    ["Shands Teaching Hospital and Clinics (UF Health Shands)", "UF Shands GNV",
     "1600 SW Archer Rd, Gainesville, FL 32610", "Gainesville", "FL", "Alachua County, FL",
     29.6400, -82.3443, "(352) 265-0111", 1162, "University of Florida Health", "Vizient",
     "CSC", 850,
     "(352) 265-0200", "Joint Commission", True, True,
     950000, 152000, 32.1, 78.6,
     False, True, True, True, True, True, False,
     2.8, 26600, 78, 50,
     True, True, "Academic flagship — full neurosurgery & NIR fellowship programs"],

    # 5 - 00004849 SHANDS JACKSONVILLE MEDICAL CENTER (UF Health Jacksonville)
    ["Shands Jacksonville Medical Center (UF Health Jax)", "UF Health Jax",
     "655 W 8th St, Jacksonville, FL 32209", "Jacksonville", "FL", "Duval County, FL",
     30.3554, -81.6721, "(904) 244-0411", 695, "University of Florida Health", "Vizient",
     "CSC", 750,
     "(904) 244-0500", "Joint Commission", True, True,
     1400000, 238000, 36.8, 76.9,
     False, True, True, True, True, True, False,
     2.9, 40600, 68, 40,
     True, True, "Academic medical center with active NIR fellowship"],

    # 6 - 00002894 SAINT JOSEPHS HOSPITAL (Tampa)
    ["Saint Joseph's Hospital", "St. Joseph's Tampa",
     "3001 W Dr Martin Luther King Jr Blvd, Tampa, FL 33607", "Tampa", "FL", "Hillsborough County, FL",
     27.9632, -82.4874, "(813) 870-4000", 897, "BayCare Health System", "Premier",
     "CSC", 710,
     "(813) 870-4100", "Joint Commission", True, True,
     1500000, 255000, 36.2, 78.0,
     False, True, True, True, True, True, False,
     3.1, 46500, 74, 45,
     True, True, "Major Tampa Bay CSC — Pepin Heart & Vascular Institute"],

    # 7 - 00005890 SAINT VINCENTS MEDICAL CENTER RIVERSIDE (Jacksonville)
    ["Saint Vincent's Medical Center Riverside", "St. Vincent's Jax",
     "1 Shircliff Way, Jacksonville, FL 32204", "Jacksonville", "FL", "Duval County, FL",
     30.3136, -81.6892, "(904) 308-7300", 528, "Ascension Health", "HPG",
     "PSC", 340,
     "", "DNV", False, False,
     850000, 144500, 39.2, 78.1,
     False, True, False, True, True, True, False,
     2.5, 21250, 55, 30,
     False, True, "Growing stroke program under new medical director"],

    # 8 - 00006101 NORTH FLORIDA REGIONAL MEDICAL CENTER (Gainesville)
    ["North Florida Regional Medical Center", "N FL Regional",
     "6500 W Newberry Rd, Gainesville, FL 32605", "Gainesville", "FL", "Alachua County, FL",
     29.6661, -82.4146, "(352) 333-4000", 510, "HCA Healthcare", "HealthTrust",
     "PSC", 290,
     "", "Joint Commission", False, False,
     480000, 81600, 34.5, 77.9,
     False, True, False, True, True, True, False,
     2.2, 10560, 48, 25,
     False, True, "HCA facility — competes with UF Shands locally"],

    # 9 - 00004852 TALLAHASSEE MEMORIAL HEALTHCARE INC
    ["Tallahassee Memorial HealthCare", "TMH",
     "1300 Miccosukee Rd, Tallahassee, FL 32308", "Tallahassee", "FL", "Leon County, FL",
     30.4408, -84.2430, "(850) 431-1155", 772, "Tallahassee Memorial HealthCare", "Premier",
     "CSC", 520,
     "(850) 431-1200", "Joint Commission", True, False,
     550000, 82500, 33.8, 77.5,
     False, True, True, True, True, True, False,
     2.4, 13200, 58, 35,
     True, True, "Only CSC between Jacksonville and Pensacola"],

    # 10 - 00004872 SHANDS TEACHING HOSPITAL (likely Shands at AGH or another campus)
    ["Shands Teaching Hospital (UF Health Rehab)", "UF Shands Rehab",
     "101 Newell Dr, Gainesville, FL 32611", "Gainesville", "FL", "Alachua County, FL",
     29.6410, -82.3468, "(352) 265-0500", 240, "University of Florida Health", "Vizient",
     "PSC", 180,
     "", "Joint Commission", False, False,
     450000, 72000, 32.1, 78.6,
     False, True, False, True, True, True, True,
     2.0, 9000, 38, 20,
     False, True, "Satellite campus of UF Health — spoke to main Shands CSC"],

    # 11 - 00004885 HCA LAKE MONROE (Central Florida Regional Hospital, Sanford)
    ["HCA Lake Monroe (Central Florida Regional)", "HCA Lake Monroe",
     "1401 W Seminole Blvd, Sanford, FL 32771", "Sanford", "FL", "Seminole County, FL",
     28.8120, -81.2804, "(407) 321-4500", 226, "HCA Healthcare", "HealthTrust",
     "PSC", 160,
     "", "Joint Commission", False, False,
     350000, 63000, 39.5, 78.2,
     False, True, False, True, False, True, True,
     2.3, 8050, 35, 15,
     False, True, ""],

    # 12 - 00006121 HCA FLORIDA LAWNWOOD HOSPITAL (Fort Pierce)
    ["HCA Florida Lawnwood Hospital", "HCA Lawnwood",
     "1700 S 23rd St, Fort Pierce, FL 34950", "Fort Pierce", "FL", "St. Lucie County, FL",
     27.4355, -80.3453, "(772) 461-4000", 380, "HCA Healthcare", "HealthTrust",
     "PSC", 250,
     "", "Joint Commission", False, False,
     420000, 88200, 42.7, 77.0,
     False, True, False, True, True, True, False,
     2.6, 10920, 45, 25,
     True, True, "Treasure Coast stroke center — high 65+ population"],

    # 13 - 00004856 BAY MEDICAL SACRED HEART HEALTH SYSTEM (Panama City)
    ["Bay Medical Sacred Heart Health System", "Bay Medical PC",
     "615 N Bonita Ave, Panama City, FL 32401", "Panama City", "FL", "Bay County, FL",
     30.1697, -85.6636, "(850) 769-1511", 323, "Ascension/Sacred Heart", "HPG",
     "PSC", 190,
     "", "DNV", False, False,
     280000, 47600, 39.8, 76.3,
     False, True, False, True, True, True, False,
     1.9, 5320, 38, 20,
     True, True, "Panhandle region — nearest CSC is TMH (100+ mi)"],

    # 14 - 00063314 MAYO CLINIC (Jacksonville)
    ["Mayo Clinic Jacksonville", "Mayo Jax",
     "4500 San Pablo Rd S, Jacksonville, FL 32224", "Jacksonville", "FL", "Duval County, FL",
     30.2617, -81.4408, "(904) 953-2000", 304, "Mayo Clinic Health System", "Mayo GPO",
     "CSC", 420,
     "(904) 953-2100", "Joint Commission", True, True,
     900000, 171000, 42.1, 81.2,
     False, True, True, True, True, True, False,
     3.8, 34200, 85, 50,
     True, True, "Research-intensive — leading clinical trials enrollment"],

    # 15 - 01413658 UF HEALTH HEART AND VASCULAR AND NEUROMEDICINE HOSPITALS
    ["UF Health Heart & Vascular and Neuromedicine Hospitals", "UF Neuro Hospital",
     "1505 SW Archer Rd, Gainesville, FL 32608", "Gainesville", "FL", "Alachua County, FL",
     29.6375, -82.3500, "(352) 265-8000", 216, "University of Florida Health", "Vizient",
     "CSC", 480,
     "(352) 265-8100", "Joint Commission", True, True,
     600000, 96000, 32.1, 78.6,
     False, True, True, True, True, True, False,
     2.6, 15600, 80, 40,
     True, True, "New dedicated neurovascular hospital — state-of-the-art hybrid ORs"],

    # 16 - 00004111 CANDLER GENERAL HOSPITAL INC (Savannah)
    ["Candler General Hospital", "Candler Savannah",
     "5353 Reynolds St, Savannah, GA 31405", "Savannah", "GA", "Chatham County, GA",
     32.0098, -81.1179, "(912) 819-6000", 384, "St. Joseph's/Candler Health System", "Premier",
     "PSC", 220,
     "", "Joint Commission", False, False,
     380000, 64600, 37.2, 75.8,
     False, True, False, True, True, True, False,
     2.3, 8740, 42, 20,
     False, True, "Part of St. Joseph's/Candler system in Savannah"],

    # 17 - 00007912 MALCOM RANDALL VA MEDICAL CENTER (Gainesville)
    ["Malcom Randall VA Medical Center", "VA Gainesville",
     "1601 SW Archer Rd, Gainesville, FL 32608", "Gainesville", "FL", "Alachua County, FL",
     29.6395, -82.3455, "(352) 376-1611", 421, "US Dept of Veterans Affairs", "VA FSS",
     "PSC", 310,
     "", "Joint Commission", False, False,
     620000, 167400, 62.5, 74.8,
     False, True, False, True, True, True, True,
     4.2, 26040, 52, 30,
     True, True, "VA facility — very high 65+ population, transfers complex cases to UF Shands"],

    # 18 - 00004837 HALIFAX HOSPITAL MEDICAL CENTER (Daytona Beach)
    ["Halifax Hospital Medical Center (Halifax Health)", "Halifax Health",
     "303 N Clyde Morris Blvd, Daytona Beach, FL 32114", "Daytona Beach", "FL", "Volusia County, FL",
     29.2184, -81.0484, "(386) 254-4000", 678, "Halifax Health", "Vizient",
     "CSC", 490,
     "(386) 254-4100", "Joint Commission", True, False,
     580000, 110200, 43.6, 77.1,
     False, True, True, True, True, True, False,
     3.0, 17400, 60, 35,
     True, True, "Central FL east coast CSC — high retiree catchment"],

    # 19 - 00002906 MEMORIAL HOSPITAL JACKSONVILLE
    ["Memorial Hospital Jacksonville", "Memorial Jax",
     "3625 University Blvd S, Jacksonville, FL 32216", "Jacksonville", "FL", "Duval County, FL",
     30.2843, -81.6021, "(904) 702-6111", 453, "HCA Healthcare", "HealthTrust",
     "PSC", 280,
     "", "Joint Commission", False, False,
     700000, 119000, 38.4, 77.8,
     False, True, False, True, True, True, False,
     2.1, 14700, 48, 25,
     False, True, ""],

    # 20 - 00006091 ORANGE PARK MEDICAL CENTER
    ["Orange Park Medical Center", "Orange Park",
     "2001 Kingsley Ave, Orange Park, FL 32073", "Orange Park", "FL", "Clay County, FL",
     30.1665, -81.7178, "(904) 639-8500", 317, "HCA Healthcare", "HealthTrust",
     "TSC", 180,
     "", "DNV", False, False,
     280000, 47600, 36.9, 77.2,
     False, True, False, True, False, True, True,
     1.8, 5040, 35, 15,
     False, True, ""],

    # 21 - 00028798 ADVENTHEALTH ORLANDO
    ["AdventHealth Orlando", "AdventHealth ORL",
     "601 E Rollins St, Orlando, FL 32803", "Orlando", "FL", "Orange County, FL",
     28.5727, -81.3627, "(407) 303-5600", 2856, "AdventHealth", "Premier",
     "CSC", 980,
     "(407) 303-5800", "Joint Commission", True, True,
     2100000, 336000, 35.8, 79.1,
     False, True, True, True, True, True, False,
     3.4, 71400, 82, 50,
     True, True, "Largest hospital in territory — 2,856 beds, massive catchment"],
]

for row_idx, row_data in enumerate(hospitals, 2):
    for col_idx, value in enumerate(row_data, 1):
        ws.cell(row=row_idx, column=col_idx, value=value)

# Column widths
ws.column_dimensions['A'].width = 50
ws.column_dimensions['B'].width = 22
ws.column_dimensions['C'].width = 45
ws.column_dimensions['D'].width = 18
ws.column_dimensions['E'].width = 6
ws.column_dimensions['F'].width = 24
for c in ['G','H']: ws.column_dimensions[c].width = 10
ws.column_dimensions['I'].width = 16
ws.column_dimensions['J'].width = 12
ws.column_dimensions['K'].width = 34
ws.column_dimensions['L'].width = 14
ws.column_dimensions['M'].width = 18
ws.column_dimensions['N'].width = 14

wb.save("/sessions/confident-beautiful-turing/jnj-territory-maps/sample_hospitals.xlsx")
print(f"Created sample_hospitals.xlsx with {len(hospitals)} hospitals")
