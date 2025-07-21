# test_logic.py
import os
from image_utils import allowed_file
from config import UPLOAD_FOLDER

def is_drop_test(title):
    return "drop test" in title['full'].lower() or "drop test" in title.get('short', '').lower()

def is_impact_test(title):
    # "impact" trong full hoặc short, không phân biệt hoa/thường
    return "impact" in title.get('full', '').lower() or "impact" in title.get('short', '').lower()

def is_rotational_test(title):
    return "rotational" in title.get('full', '').lower() or "rotational" in title.get('short', '').lower()

def load_group_notes(file_path):
    notes = {}
    if os.path.exists(file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                if line.strip().startswith("Mục "):
                    try:
                        muc, val = line.strip().split(":", 1)
                        key = muc.strip().replace("Mục ", "")
                        notes[key] = val.strip()
                    except:
                        pass
    return notes

def get_group_test_status(report, group, test_key):
    report_folder = os.path.join(UPLOAD_FOLDER, str(report))
    status_file = os.path.join(report_folder, f"status_{group}_{test_key}.txt")
    comment_file = os.path.join(report_folder, f"comment_{group}_{test_key}.txt")
    status = comment = None
    has_img = False
    first_img = None
    if os.path.exists(report_folder):
        for f in sorted(os.listdir(report_folder)):
            if allowed_file(f) and f.startswith(f"test_{group}_{test_key}_"):
                has_img = True
                first_img = f"/images/{report}/{f}"
                break
    if os.path.exists(status_file):
        with open(status_file, 'r', encoding='utf-8') as f:
            status = f.read().strip()
    if os.path.exists(comment_file):
        with open(comment_file, 'r', encoding='utf-8') as f:
            comment = f.read().strip()
    return {'status': status, 'comment': comment, 'has_img': has_img, 'first_img': first_img}

BAN_US_TEST_TITLES = {
    "muc4.2": {
        "full": "Mục 4.2: Stability with extendible elements open test",
        "short": "Stability with extendible elements open test",
        "img": ["/static/images/buoc 3 ban us.jpg"]
    },
    "muc4": {
        "full": "Mục 4: Stability test",
        "short": "Stability test",
        "img": ["/static/images/buoc 4 ban us.png"]
    },
    "muc4.4": {
        "full": "Mục 4.4: Horizontal Stability Test for Desks and Tables with Casters",
        "short": "Horizontal Stability Test for Desks and Tables with Casters",
        "img": ["/static/images/buoc 5 ban us.png"]
    },
    "muc4.5": {
        "full": "Mục 4.5: Horizontal Stability Test for Keyboard/Laptop Tables (with and without casters)",
        "short": "Horizontal Stability Test for Keyboard/Laptop Tables (with and without casters)",
        "img": ["/static/images/buoc 6 ban us.jpg"]
    },
    "muc5.2": {
        "full": "Mục 5.2: Concentrated load test – Functional",
        "short": "Concentrated load test – Functional",
        "img": ["/static/images/buoc 7 ban us.jpg"]
    },
    "muc5.3": {
        "full": "Mục 5.3: Distributed load test - Functional",
        "short": "Distributed load test - Functional",
        "img": ["/static/images/buoc 8 ban us.png"]
    },
    "muc5.4": {
        "full": "Mục 5.4: Concentrated load test - Proof",
        "short": "Concentrated load test - Proof",
        "img": ["/static/images/buoc 9 ban us.png"]
    },
    "muc5.5": {
        "full": "Mục 5.5: Distributed load test - Proof",
        "short": "Distributed load test - Proof",
        "img": ["/static/images/buoc 10 ban us.png"]
    },
    "muc5.6": {
        "full": "Mục 5.6: Transaction Surface Torsion Load Test",
        "short": "Transaction Surface Torsion Load Test",
        "img": ["/static/images/buoc 11 ban us.png"]
    },
    "muc5.7": {
        "full": "Mục 5.7: Extendible element static load test",
        "short": "Extendible element static load test",
        "img": ["/static/images/buoc 12 ban us.png"]
    },
    "muc7": {
        "full": "Mục 7: Desk or table unit drop test",
        "short": "Desk or table unit drop test",
        "img": ["/static/images/buoc 13 ban us.png"]
    },
    "muc8.2.1": {
        "full": "Mục 8.2.1: Leg strength test – Set up",
        "short": "Leg strength test – Set up",
        "img": ["/static/images/buoc 14 ban us.png"]
    },
    "muc8.2.2": {
        "full": "Mục 8.2.2: Standard - Leg strength test - Functional load",
        "short": "Leg strength test - Functional load",
        "img": ["/static/images/buoc 15 ban us.png"]
    },
    "muc8.2.3": {
        "full": "Mục 8.2.3: Standard - Leg strength test – Proof load",
        "short": "Leg strength test – Proof load",
        "img": ["/static/images/buoc 16 ban us.png"]
    },
    "muc8.3.2": {
        "full": "Mục 8.3.2: Alternate - Leg Strength Test - Functional",
        "short": "Alternate - Leg Strength Test - Functional",
        "img": ["/static/images/buoc 17 ban us.png"]
    },
    "muc8.3.3": {
        "full": "Mục 8.3.3: Alternate - Leg Strength Test - Proof",
        "short": "Alternate - Leg Strength Test - Proof",
        "img": ["/static/images/buoc 18 ban us.png"]
    },
    "muc11": {
        "full": "Mục 11: Extendible Element Retention Impact and Durability (Out Stop) Tests",
        "short": "Extendible Element Retention Impact and Durability Tests",
        "img": ["/static/images/buoc 19 ban us.png"]
    },
    "muc12": {
        "full": "Mục 12: Extendible element rebound test",
        "short": "Extendible element rebound test",
        "img": ["/static/images/buoc 20 ban us.png"]
    },
    "muc17": {
        "full": "Mục 17: Strength Test for Vertically Hinged Doors, Bi-fold Doors and Vertically Receding Doors",
        "short": "Strength Test for Vertically Hinged/Bi-fold/Receding Doors",
        "img": ["/static/images/buoc 21 ban us.png"]
    },
    "muc19": {
        "full": "Mục 19: Pull force test",
        "short": "Pull force test",
        "img": ["/static/images/buoc 22 ban us.png"]
    }
}

BAN_EU_TEST_TITLES = {
    "muc5.5.2_7.2.2_7.2.3": {
        "full": "Mục (5.5.2 - 12521:2023) + (7.2.2, 7.2.3 - 1730:2012): Stability under vertical load",
        "short": "Stability under vertical load",
        "img": ["/static/images/buoc 3 ban eu.png"]
    },
    "muc5.5.3_7.3": {
        "full": "Mục (5.5.3 - 12521:2023) + (7.3 - 1730:2012): Stability for tables with extension elements",
        "short": "Stability for tables with extension elements",
        "img": ["/static/images/buoc 4 ban eu.png"]
    },
    "muc5.6_6.2": {
        "full": "Mục (5.6 - 12521:2023) + (6.2 - 1730:2012): Horizontal static load test",
        "short": "Horizontal static load test",
        "img": ["/static/images/buoc 5 ban eu.png"]
    },
    "muc5.6_6.3.1": {
        "full": "Mục (5.6 - 12521:2023) + (6.3.1 - 1730:2012): Vertical static load on main surface tests",
        "short": "Vertical static load on main surface tests",
        "img": ["/static/images/buoc 6 ban eu.png"]
    },
    "muc5.6_6.3.2": {
        "full": "Mục (5.6 - 12521:2023) + (6.3.2 - 1730:2012): Additional vertical static load test where the main surface has a length > 1 600 mm",
        "short": "Additional vertical static load test (>1600mm)",
        "img": ["/static/images/buoc 7 ban eu.png"]
    },
    "muc5.6_6.3.3": {
        "full": "Mục (5.6 - 12521:2023) + (6.3.3 - 1730:2012): Vertical static load on ancillary surface",
        "short": "Vertical static load on ancillary surface",
        "img": ["/static/images/buoc 8 ban eu.png"]
    },
    "muc6.6.1_6.6.3": {
        "full": "Mục (6.6.1 & 6.6.3 - 1730:2012): Vertical impact test without glass",
        "short": "Vertical impact test without glass",
        "img": ["/static/images/buoc 9 ban eu.png"]
    },
    "muc6.6.1_6.6.2": {
        "full": "Mục (6.6.1 & 6.6.2 - 1730:2012): Vertical impact test with glass",
        "short": "Vertical impact test with glass",
        "img": ["/static/images/buoc 10 ban eu.png"]
    },
    "muc6.9": {
        "full": "Mục 6.9: Drop test (Mod) ANSI BIFMA X5.5",
        "short": "Drop test (Mod) ANSI BIFMA X5.5",
        "img": ["/static/images/buoc 11 ban eu.png"]
    }
}

GHE_US_TEST_TITLES = {
    "muc5.4": {
        "full": "Mục 5.4.1 + 5.4.2: Backrest strength test - horizontal – static (functional load, proof load)",
        "short": "Backrest strength test - horizontal – static (func./proof load)",
        "img": ["/static/images/buoc 2 ghe us.png"]
    },
    "muc6": {
        "full": "Mục 6: Backrest Strength Test - Vertical - Static",
        "short": "Backrest Strength Test - Vertical - Static",
        "img": ["/static/images/buoc 3 ghe us.png"]
    },
    "muc9.5": {
        "full": "Mục 9.5.1 + 9.5.2: Arm strength test - horizontal – static (functional load, proof load)",
        "short": "Arm strength test - horizontal – static (func./proof load)",
        "img": ["/static/images/buoc 4 ghe us.png"]
    },
    "muc10": {
        "full": "Mục 10: Arm strength test - vertical – static (functional load, proof load)",
        "short": "Arm strength test - vertical – static (func./proof load)",
        "img": ["/static/images/buoc 5 ghe us.png"]
    },
    "muc15": {
        "full": "Mục 15: Drop test – dynamic (functional load, proof load)",
        "short": "Drop test – dynamic (func./proof load)",
        "img": ["/static/images/buoc 6 ghe us.png"]
    },
    "muc16.3": {
        "full": "Mục 16.3: Leg strength test - front load test (functional load, proof load)",
        "short": "Leg strength test - front load (func./proof load)",
        "img": ["/static/images/buoc 7 ghe us.png"]
    },
    "muc16.4": {
        "full": "Mục 16.4: Leg strength test - side load test (functional load, proof load)",
        "short": "Leg strength test - side load (func./proof load)",
        "img": ["/static/images/buoc 8 ghe us.png"]
    },
    "muc17": {
        "full": "Mục 17: Unit drop test - dynamic",
        "short": "Unit drop test - dynamic",
        "img": ["/static/images/buoc 9 ghe us.png"]
    },
    "muc21.3": {
        "full": "Mục 21.3: Stability tests - rear stability for non-tilting units",
        "short": "Stability test - rear stability",
        "img": ["/static/images/buoc 10 ghe us.png"]
    },
    "muc21.5": {
        "full": "Mục 21.5: Forward Stability for units < 36.3 kg (80 lbs)",
        "short": "Forward Stability < 36.3kg",
        "img": ["/static/images/buoc 11 ghe us.png"]
    },
    "muc21.6": {
        "full": "Mục 21.6: Forward Stability for units ≥ 36.3 kg (80 lbs)",
        "short": "Forward Stability ≥ 36.3kg",
        "img": ["/static/images/buoc 12 ghe us.png"]
    },
    "muc27": {
        "full": "Mục 27: Footrest Static Load Test for Stools - Vertical",
        "short": "Footrest Static Load Test - Vertical",
        "img": ["/static/images/buoc 13 ghe us.png"]
    },
    "vfr_drop": {
        "full": "VFR Drop test (Mod)",
        "short": "VFR Drop test (Mod)",
        "img": ["/static/images/buoc 14 ghe us.png"]
    },
    "dist_load": {
        "full": "Distributed load capacity (Mod)",
        "short": "Distributed load capacity (Mod)",
        "img": ["/static/images/buoc 15 ghe us.png"]
    }
}

GHE_EU_TEST_TITLES = {
    "muc6.4": {
        "full": "Mục 6.4: Seat static & back static load test",
        "short": "Seat & back static load",
        "img": ["/static/images/buoc 2 ghe eu.png"]
    },
    "muc6.5": {
        "full": "Mục 6.5: Seat front edge static load",
        "short": "Seat front edge static load",
        "img": ["/static/images/buoc 3 ghe eu.png"]
    },
    "muc6.6": {
        "full": "Mục 6.6: Vertical static load on back (b)",
        "short": "Vertical static load on back",
        "img": ["/static/images/buoc 4 ghe eu.png"]
    },
    "muc6.8_6.9": {
        "full": "Mục 6.8 + 6.9: Foot rest static load test & Leg rest static load test",
        "short": "Foot rest & Leg rest static load",
        "img": ["/static/images/buoc 5 ghe eu.png"]
    },
    "muc6.10": {
        "full": "Mục 6.10: Arm rest sideways static load test",
        "short": "Arm rest sideways static load",
        "img": ["/static/images/buoc 6 ghe eu.png"]
    },
    "muc6.11": {
        "full": "Mục 6.11: Arm rest downwards static load test",
        "short": "Arm rest downwards static load",
        "img": ["/static/images/buoc 7 ghe eu.png"]
    },
    "muc6.13": {
        "full": "Mục 6.13.1 + 6.13.2: Vertical upwards static load on arm rests",
        "short": "Arm rests vertical upwards static load",
        "img": ["/static/images/buoc 8 ghe eu.png"]
    },
    "muc6.15": {
        "full": "Mục 6.15: Leg forward static load test",
        "short": "Leg forward static load",
        "img": ["/static/images/buoc 9 ghe eu.png"]
    },
    "muc6.16": {
        "full": "Mục 6.16: Leg sideways static load test",
        "short": "Leg sideways static load",
        "img": ["/static/images/buoc 10 ghe eu.png"]
    },
    "muc6.24": {
        "full": "Mục 6.24: Seat impact test",
        "short": "Seat impact test",
        "img": ["/static/images/buoc 11 ghe eu.png"]
    },
    "muc6.25": {
        "full": "Mục 6.25: Back impact test",
        "short": "Back impact test",
        "img": ["/static/images/buoc 12 ghe eu.png"]
    },
    "muc6.26": {
        "full": "Mục 6.26: Arm rest impact test",
        "short": "Arm rest impact test",
        "img": ["/static/images/buoc 13 ghe eu.png"]
    },
    "muc6.28": {
        "full": "Mục 6.28: Backward fall test",
        "short": "Backward fall test",
        "img": ["/static/images/buoc 14 ghe eu.png"]
    },
    "muc7.3.1": {
        "full": "Mục 7.3.1 (1022:2018): Forwards overturning",
        "short": "Forwards overturning",
        "img": ["/static/images/buoc 15 ghe eu.png"]
    },
    "muc7.3.3": {
        "full": "Mục 7.3.3 (1022:2018): Corner stability test",
        "short": "Corner stability test",
        "img": ["/static/images/buoc 16 ghe eu.png"]
    },
    "muc7.3.4": {
        "full": "Mục 7.3.4 (1022:2018): Sideways overturning, all seating without arm rest",
        "short": "Sideways overturning (no arm rest)",
        "img": ["/static/images/buoc 17 ghe eu.png"]
    },
    "muc7.3.5.2": {
        "full": "Mục 7.3.5.2 (1022:2018): Seating with arm rests",
        "short": "Overturning: seating with arm rests",
        "img": ["/static/images/buoc 18 ghe eu.png"]
    },
    "muc7.3.6": {
        "full": "Mục 7.3.6 (1022:2018): Rearwards overturning all seating with backrest",
        "short": "Rearwards overturning",
        "img": ["/static/images/buoc 19 ghe eu.png"]
    }
}

TU_US_TEST_TITLES = {
    "muc3": {
        "full": "Mục 4.2: Concentrated functional load test",
        "short": "Concentrated functional load test",
        "img": ["/static/images/buoc 3 tu us.png"]
    },
    "muc4": {
        "full": "Mục 4.3: Distributed functional load test",
        "short": "Distributed functional load test",
        "img": ["/static/images/buoc 4 tu us.png"]
    },
    "muc5": {
        "full": "Mục 4.4: Concentrated proof load test",
        "short": "Concentrated proof load test",
        "img": ["/static/images/buoc 5 tu us.png"]
    },
    "muc6": {
        "full": "Mục 4.5: Distributed Proof Load Test",
        "short": "Distributed Proof Load Test",
        "img": ["/static/images/buoc 6 tu us.png"]
    },
    "muc7": {
        "full": "Mục 4.6.2: Extendible Element Functional Load Test",
        "short": "Extendible Element Functional Load Test",
        "img": ["/static/images/buoc 7 tu us.png"]
    },
    "muc8": {
        "full": "Mục 4.6.3: Extendible Element Proof Load Test",
        "short": "Extendible Element Proof Load Test",
        "img": ["/static/images/buoc 8 tu us.png"]
    },
    "muc9": {
        "full": "Mục 5.4: Leg/Glide Assembly Strength Test (Functional và Proof) Sec 5.3 và 5.6",
        "short": "Leg/Glide Assembly Strength Test",
        "img": ["/static/images/buoc 9 tu us.png"]
    },
    "muc10": {
        "full": "Mục 6: Racking Resistance Test",
        "short": "Racking Resistance Test",
        "img": ["/static/images/buoc 10 tu us.png"]
    },
    "muc11": {
        "full": "Mục 7.2: Drop Test - Dynamic - for Units with Seat Surfaces",
        "short": "Drop Test - Dynamic",
        "img": ["/static/images/buoc 11 tu us.png"]
    },
    "muc12": {
        "full": "Mục 8.1: Separation Test for Tall Storage Units with Vertically Attached or Stackable Components",
        "short": "Separation Test (Tall Storage)",
        "img": ["/static/images/buoc 12 tu us.png"]
    },
    "muc13": {
        "full": "Mục 9.2: Horizontal force stability test for storage units without extendible elements",
        "short": "Horizontal force stability (no extendible)",
        "img": ["/static/images/buoc 13 tu us.png"]
    },
    "muc14": {
        "full": "Mục 9.3: Stability test for type I units with at least one extendible elements",
        "short": "Stability test type I",
        "img": ["/static/images/buoc 14 tu us.png"]
    },
    "muc15": {
        "full": "Mục 9.4: Stability Test for Type I Storage Units with Multiple Extendible Elements",
        "short": "Stability Test Type I Multiple",
        "img": ["/static/images/buoc 15 tu us.png"]
    },
    "muc16": {
        "full": "Mục 9.5: Stability Test for Type II Storage Units with Extendible Elements",
        "short": "Stability Test Type II",
        "img": ["/static/images/buoc 16 tu us.png"]
    },
    "muc17": {
        "full": "Mục 9.6: Vertical Force Stability Test for Storage Units",
        "short": "Vertical Force Stability",
        "img": ["/static/images/buoc 17 tu us.png"]
    },
    "muc18": {
        "full": "Mục 9.7: Stability Test for Pedestals/Storage Units with Seat Surfaces",
        "short": "Stability Pedestals/Seat Surfaces",
        "img": ["/static/images/buoc 18 tu us.png"]
    },
    "muc19": {
        "full": "Mục 9.9: Extendible element rebound test",
        "short": "Extendible element rebound test",
        "img": ["/static/images/buoc 19 tu us.png"]
    },
    "muc20": {
        "full": "Mục 12: Extendible element rebound test",
        "short": "Extendible element rebound test",
        "img": ["/static/images/buoc 20 tu us.png"]
    },
    "muc21": {
        "full": "Mục 13: Extendible Element Retention Impact and Durability (Out Stop) Tests",
        "short": "Retention Impact/Durability (Out Stop)",
        "img": ["/static/images/buoc 21 tu us.png"]
    },
    "muc22": {
        "full": "Mục 17.2: Strength Test for Vertically Hinged Doors, Bi-fold Doors and Vertically Receding Doors",
        "short": "Strength Test Hinged/Bi-fold/Receding Doors",
        "img": ["/static/images/buoc 22 tu us.png"]
    },
    "muc23": {
        "full": "Mục 17.3: Hinge Override Test for Vertically Hinged Doors (See Figure 17b)",
        "short": "Hinge Override Test Hinged Doors",
        "img": ["/static/images/buoc 23 tu us.png"]
    },
    "muc24": {
        "full": "Mục 20: Pull force test",
        "short": "Pull force test",
        "img": ["/static/images/buoc 24 tu us.png"]
    }
}

TU_EU_TEST_TITLES = {
    "muc3": {
        "full": "Mục 5.2.5: Extension elements (EN14749:2016)",
        "short": "Extension elements",
        "img": ["/static/images/buoc 3 tu eu.png"]
    },
    "muc4": {
        "full": "Mục 5.3.2.1: Shelf retention - vertical downward (EN14749:2016); Mục 6.1.3 (EN16122:2012)",
        "short": "Shelf retention - vertical downward",
        "img": ["/static/images/buoc 4 tu eu.png"]
    },
    "muc5": {
        "full": "Mục 5.3.2.2: Shelf retention - horizontal outward (EN14749:2016); Mục 6.1.2 (EN16122:2012)",
        "short": "Shelf retention - horizontal outward",
        "img": ["/static/images/buoc 5 tu eu.png"]
    },
    "muc6": {
        "full": "Mục 5.3.3: Shelf supports (EN14749:2016); Mục 6.1.5 (EN16122:2012)",
        "short": "Shelf supports",
        "img": ["/static/images/buoc 6 tu eu.png"]
    },
    "muc7": {
        "full": "Mục 5.3.5.1: Vertical load of pivoted doors (EN14749:2016); Mục 7.1.1, 7.1.2 (EN16122:2012)",
        "short": "Vertical load of pivoted doors",
        "img": ["/static/images/buoc 7 tu eu.png"]
    },
    "muc8": {
        "full": "Mục 5.3.5.2: Horizontal load on pivoted doors (EN14749:2016); Mục 7.1.1, 7.1.3 (EN16122:2012)",
        "short": "Horizontal load on pivoted doors",
        "img": ["/static/images/buoc 8 tu eu.png"]
    },
    "muc9": {
        "full": "Mục 5.3.7.2: Strength test of extension elements (EN14749:2016); Mục 7.5.1, 7.5.2 (EN16122:2012)",
        "short": "Strength test of extension elements",
        "img": ["/static/images/buoc 9 tu eu.png"]
    },
    "muc10": {
        "full": "Mục 5.3.7.1: Slam open of extension elements (EN14749:2016); Mục 7.5.1, 7.5.4 (EN16122:2012)",
        "short": "Slam open of extension elements",
        "img": ["/static/images/buoc 10 tu eu.png"]
    },
    "muc11": {
        "full": "Mục 8.3: Drop test for trays (EN16122:2012)",
        "short": "Drop test for trays",
        "img": ["/static/images/buoc 11 tu eu.png"]
    },
    "muc12": {
        "full": "Mục 5.4.1.1: Stability (EN14749:2016); Mục 11.2.1, 11.2.2 (EN16122:2012)",
        "short": "Stability",
        "img": ["/static/images/buoc 12 tu eu.png"]
    },
    "muc13": {
        "full": "Mục 5.4.1.3: Stability - unloaded - all doors, extension elements and flaps open (EN14749:2016); Mục 11.4.1 (EN16122:2012)",
        "short": "Stability - unloaded (all open)",
        "img": ["/static/images/buoc 13 tu eu.png"]
    },
    "muc14": {
        "full": "Mục 5.4.1.4: Stability - unloaded - with overturning load (EN14749:2016); Mục 11.4.2 (EN16122:2012)",
        "short": "Stability - unloaded (overturning load)",
        "img": ["/static/images/buoc 14 tu eu.png"]
    },
    "muc15": {
        "full": "Mục 5.4.1.5: Stability - loaded - with overturning load (EN14749:2016); Mục 11.4.3 (EN16122:2012)",
        "short": "Stability - loaded (overturning load)",
        "img": ["/static/images/buoc 15 tu eu.png"]
    }
}

GIUONG_TEST_TITLES = {
    "muc2": {
        "full": "Distributed load capacity",
        "short": "Distributed load",
        "img": ["/static/images/buoc 2 giuong.png"]
    },
    "muc3": {
        "full": "Impact durability",
        "short": "Impact durability",
        "img": ["/static/images/buoc 3 giuong.png"]
    },
    "muc4": {
        "full": "Bed - vertical static load test (EN 1725-98 Cl. 7.6 Mod)",
        "short": "Vertical static load test (center)",
        "img": ["/static/images/buoc 4 giuong.png"]
    },
    "muc5": {
        "full": "Bed - vertical static load test of the edge of the bed (EN 1725-98 Cl. 7.7 Mod)",
        "short": "Vertical static load test (edge)",
        "img": ["/static/images/buoc 5 giuong.png"]
    },
    "muc6": {
        "full": "Headboard Pull test",
        "short": "Headboard Pull test",
        "img": ["/static/images/buoc 6 giuong.png"]
    },
    "muc7": {
        "full": "End drop test (ANSI/BIFMA X5.5-21 Sec. 7 Mod)",
        "short": "End drop test",
        "img": ["/static/images/buoc 7 giuong.png"]
    },
    "muc8": {
        "full": "Canopy frame static load test (khung màn che - không có rèm)",
        "short": "Canopy frame static load (no curtain)",
        "img": ["/static/images/buoc 8 giuong.png"]
    },
    "muc9": {
        "full": "Canopy frame static load test (khung màn che - có rèm)",
        "short": "Canopy frame static load (with curtain)",
        "img": ["/static/images/buoc 9 giuong.png"]
    },
    "muc10": {
        "full": "Hanging strength",
        "short": "Hanging strength",
        "img": ["/static/images/buoc 10 giuong.png"]
    }
}

EXTRA_TEST = {
    "test_extra": {
        "full": "Mục bổ sung: Kiểm tra đặc biệt",
        "short": "Kiểm tra đặc biệt",
    }
}

INDOOR_CHUYEN_TEST_TITLES = {
    "pencil": {
        "full": "1: Pencil hardness test",
        "short": "Pencil",
        "img": ["/static/images/buoc 1 indoor.png"]  # 1 ảnh
    },
    "adhesion": {
        "full": "2: Adhesion test",
        "short": "Adhesion",
        "img": [
            "/static/images/buoc 3.1 indoor.png",  # Ảnh 1
            "/static/images/buoc 3.2 indoor.png"   # Ảnh 2
        ]
    },
    "standing_water": {
        "full": "3: Standing water test",
        "short": "Standing water",
        "img": ["/static/images/buoc 5 indoor.png"]
    },
    "hot_cold": {
        "full": "4: Hot and cold cycle test",
        "short": "Hot and cold",
        "img": ["/static/images/buoc 6 indoor.png"]
    },
    "impact": {
        "full": "5: Impact resistance test",
        "short": "Impact",
        "img": ["/static/images/buoc 7 indoor.png"]
    }
}

INDOOR_THUONG_TEST_TITLES = {
    "pencil": {
        "full": "1: Pencil hardness test",
        "short": "Pencil",
        "img": ["/static/images/buoc 1 indoor.png"]  # 1 ảnh
    },
    "adhesion": {
        "full": "2: Adhesion test",
        "short": "Adhesion",
        "img": [
            "/static/images/buoc 3.1 indoor.png",  # Ảnh 1
            "/static/images/buoc 3.2 indoor.png"   # Ảnh 2
        ]
    },
    "standing_water": {
        "full": "3: Standing water test",
        "short": "Standing water",
        "img": ["/static/images/buoc 5 indoor.png"]
    },
    "hot_cold": {
        "full": "4: Hot and cold cycle test",
        "short": "Hot and cold",
        "img": ["/static/images/buoc 6 indoor.png"]
    },
    "impact": {
        "full": "5: Impact resistance test",
        "short": "Impact",
        "img": ["/static/images/buoc 7 indoor.png"]
    },
    "stain": {
        "full": "6: Stain test",
        "short": "Stain",
        "img": ["/static/images/buoc 8 indoor.png"]
    },
    "solvent": {
        "full": "7: Solvent test",
        "short": "Solvent",
        "img": ["/static/images/buoc 15 indoor.png"]
    }
}

INDOOR_STONE_TEST_TITLES = {
    "pencil": {
        "full": "1: Pencil hardness test",
        "short": "Pencil",
        "img": ["/static/images/buoc 1 indoor.png"]  # 1 ảnh
    },
    "adhesion": {
        "full": "2: Adhesion test",
        "short": "Adhesion",
        "img": [
            "/static/images/buoc 3.1 indoor.png",  # Ảnh 1
            "/static/images/buoc 3.2 indoor.png"   # Ảnh 2
        ]
    },
    "hot_cold": {
        "full": "3: Hot and cold cycle test",
        "short": "Hot and cold",
        "img": ["/static/images/buoc 6 indoor.png"]
    },
    "impact": {
        "full": "4: Impact resistance test",
        "short": "Impact",
        "img": ["/static/images/buoc 7 indoor.png"]
    },
    "stain": {
        "full": "5: Stain test",
        "short": "Stain",
        "img": ["/static/images/buoc 8 indoor.png"]
    },
}

INDOOR_METAL_TEST_TITLES = {
    "pencil": {
        "full": "1: Pencil hardness test",
        "short": "Pencil",
        "img": ["/static/images/buoc 1 indoor.png"]  # 1 ảnh
    },
    "adhesion": {
        "full": "2: Adhesion test",
        "short": "Adhesion",
        "img": [
            "/static/images/buoc 3.1 indoor.png",  # Ảnh 1
            "/static/images/buoc 3.2 indoor.png"   # Ảnh 2
        ]
    },
    "corrosion": {
        "full": "3:Corrosion test",
        "short": "Corrosion",
        "img": ["/static/images/buoc 4 indoor.png"]
    },
    "hot_cold": {
        "full": "4: Hot and cold cycle test",
        "short": "Hot and cold",
        "img": ["/static/images/buoc 6 indoor.png"]
    },
    "impact": {
        "full": "5: Impact resistance test",
        "short": "Impact",
        "img": ["/static/images/buoc 7 indoor.png"]
    },
    "solvent": {
        "full": "6: Solvent test",
        "short": "Solvent",
        "img": ["/static/images/buoc 15 indoor.png"]
    }
}

OUTDOOR_FINISHING_TEST_TITLES = {
    "muc1": {
        "full": "1: Corrosion test 5% - Áp dụng kim loại",
        "short": "Corrosion test 5% (kim loại)",
        "img": ["/static/images/buoc 1 outdoor.jpg"]
    },
    "muc2": {
        "full": "2: Stain resistance - Áp dụng đá",
        "short": "Stain resistance (đá)",
        "img": ["/static/images/buoc 2 outdoor.jpg"]
    },
    "muc3": {
        "full": "3: Before adhesion",
        "short": "Adhesion before",
        "img": ["/static/images/buoc 3 outdoor.jpg"]
    },
    "muc4": {
        "full": "4: Hydrothermal",
        "short": "Hydrothermal",
        "img": ["/static/images/buoc 4 outdoor.jpg"]
    },
    "muc5": {
        "full": "5: After adhesion",
        "short": "After adhesion",
        "img": ["/static/images/buoc 5 outdoor.jpg"]
    }
}

TRANSIT_2C_NP_TEST_TITLES = {
    "step1": {
        "full": "Bước 1: Kiểm tra thông tin sản phẩm",
        "short": "Thông tin sản phẩm",
        "img": ["/static/images/buoc 1 2C std.png"]
    },
    "step2": {
        "full": "Bước 2: Identification of Faces, Edges and Corners",
        "short": "Nhận diện mặt, cạnh, góc",
        "img": ["/static/images/buoc 2 2C std.png"]
    },
    "step3": {
        "full": "Bước 3: Before Vibration Under Dynamic Load",
        "short": "Chuẩn bị rung động (dynamic load)",
        "img": ["/static/images/buoc 3 2C std.png"]
    },
    "step5": {
        "full": "Bước 5: Vibration Under Dynamic Load",
        "short": "Rung động (dynamic load)",
        "img": ["/static/images/buoc 5 2C std.png"]
    },
    "step6": {
        "full": "Bước 6: Drop test",
        "short": "Thả rơi",
        "img": ["/static/images/buoc 6 2C std.png"]
    },
    "step10": {
        "full": "Bước 10: Kiểm tra lại trong quá trình test",
        "short": "Kiểm tra lại sau test",
        "img": ["/static/images/buoc 10 2C std.png"]
    }
}

TRANSIT_2C_PALLET_TEST_TITLES = {
    "step1": {
        "full": "Bước 1: Kiểm tra thông tin sản phẩm",
        "short": "Thông tin sản phẩm",
        "img": ["/static/images/buoc 1 2C std.png"]
    },
    "step2": {
        "full": "Bước 2: Identification of Faces, Edges and Corners",
        "short": "Nhận diện mặt, cạnh, góc",
        "img": ["/static/images/buoc 2 2C std.png"]
    },
    "step3": {
        "full": "Bước 3: Before Vibration Under Dynamic Load",
        "short": "Chuẩn bị rung động (dynamic load)",
        "img": ["/static/images/buoc 3 2C std.png"]
    },
    "step5": {
        "full": "Bước 5: Vibration Under Dynamic Load",
        "short": "Rung động (dynamic load)",
        "img": ["/static/images/buoc 5 2C std.png"]
    },
    "step8": {
        "full": "Bước 8: Exception Two – Shock - Impact",
        "short": "Shock - Impact (Exception Two)",
        "img": ["/static/images/buoc 8 2C std.png"]
    },
    "step9": {
        "full": "Bước 9: Exception Two – Rotational Edge Drop",
        "short": "Rotational Edge Drop",
        "img": ["/static/images/buoc 9 2C std.png"]
    },
    "step10": {
        "full": "Bước 10: Kiểm tra lại trong quá trình test",
        "short": "Kiểm tra lại sau test",
        "img": ["/static/images/buoc 10 2C std.png"]
    }
}

TRANSIT_RH_NP_TEST_TITLES = {
    "step1": {
        "full": "Bước 1: Kiểm tra thông tin sản phẩm",
        "short": "Thông tin sản phẩm",
        "img": ["/static/images/buoc 1 RH all.png"]
    },
    "step2": {
        "full": "Bước 2: Compression - Top Load (ASTM D642-00 R2010)",
        "short": "Tải nén mặt trên",
        "img": ["/static/images/buoc 2 RH all.png"]
    },
    "step3": {
        "full": "Bước 3: First Impact Test Series (ASTM D5276)",
        "short": "Thả rơi lần 1",
        "img": ["/static/images/buoc 3 RH.png"]
    },
    "step4": {
        "full": "Bước 4: Loose Load Vibration (ASTM D4169-09/D999-08)",
        "short": "Rung không tải",
        "img": ["/static/images/buoc 4 RH.png"]
    },
    "step5": {
        "full": "Bước 5: Second Impact Test Series (ASTM D5276)",
        "short": "Thả rơi lần 2",
        "img": ["/static/images/buoc 5 RH.png"]
    },
    "step6": {
        "full": "Bước 6: RH Special (Dưới 9kg - Nâng lên 762mm, ném 10ft)",
        "short": "RH Special <9kg",
        "img": ["/static/images/buoc 6 RH.png"]
    },
    "step7": {
        "full": "Bước 7: RH Flat - Vật nặng 10lbs thả lên mặt lớn nhất",
        "short": "Flat - Vật nặng",
        "img": ["/static/images/buoc 7 RH.png"]
    },
    "step8": {
        "full": "Bước 8: Rotational Flat Impacts",
        "short": "Thả xoay đầu Flat",
        "img": ["/static/images/buoc 8 RH.png"]
    },
    "step9": {
        "full": "Bước 9: Rotational Edge Impacts",
        "short": "Thả xoay cạnh",
        "img": ["/static/images/buoc 9 RH.png"]
    },
    "step10": {
        "full": "Bước 10: Stability – Kiểm tra nghiêng 22°",
        "short": "Kiểm tra nghiêng 22°",
        "img": ["/static/images/buoc 10 RH.png"]
    },
    "step11": {
        "full": "Bước 11: Tip-over – Cho ngã trên nền nhà",
        "short": "Tip-over (cho ngã)",
        "img": ["/static/images/buoc 11 RH.png"]
    },
    "step12": {
        "full": "Bước 12: Flat Drop – Drop Test mặt 3 (>91kg ASTM D5276)",
        "short": "Drop test mặt 3",
        "img": ["/static/images/buoc 12 RH.png"]
    },
    "step13": {
        "full": "Bước 13: Kiểm tra lại trong quá trình test (sơn, nứt, bể...)",
        "short": "Kiểm tra lại sau test",
        "img": ["/static/images/buoc 13 RH.png"]
    }
}

TRANSIT_RH_PALLET_TEST_TITLES = {
    "step1": {
        "full": "Bước 1: Kiểm tra thông tin sản phẩm",
        "short": "Thông tin sản phẩm",
        "img": ["/static/images/buoc 1 RH all.png"]
    },
    "step2": {
        "full": "Bước 2: Compression - Top Load (ASTM D642-00 R2010)",
        "short": "Tải nén mặt trên",
        "img": ["/static/images/buoc 2 RH all.png"]
    },
    "step3": {
        "full": "Bước 3: Loose Load Vibration (ASTM D4169-09/D999-08)",
        "short": "Rung mặt 3",
        "img": ["/static/images/buoc 3 RH pal.png"]
    },
    "step4": {
        "full": "Bước 4: Stability – Kiểm tra nghiêng 22°",
        "short": "Kiểm tra nghiêng 22°",
        "img": ["/static/images/buoc 4 RH pal.png"]
    },
    "step5": {
        "full": "Bước 5: Tip-over – Cho ngã trên nền nhà (chỉ khi fail bước Stability)",
        "short": "Tip-over (cho ngã)",
        "img": ["/static/images/buoc 5 RH pal.png"]
    },
    "step6": {
        "full": "Bước 6: Rotational Edge Impacts",
        "short": "Thả xoay cạnh",
        "img": ["/static/images/buoc 6 RH pal.png"]
    },
    "step7": {
        "full": "Bước 7: Kiểm tra lại trong quá trình test (sơn, nứt, bể...)",
        "short": "Kiểm tra lại sau test",
        "img": ["/static/images/buoc 7 RH pal.png"]
    }
}

TEST_GROUP_TITLES = {
    'ban_us': BAN_US_TEST_TITLES,
    'ban_eu': BAN_EU_TEST_TITLES,
    'ghe_us': GHE_US_TEST_TITLES,
    'ghe_eu': GHE_EU_TEST_TITLES,
    'tu_us' : TU_US_TEST_TITLES,
    'tu_eu' : TU_EU_TEST_TITLES,
    'giuong': GIUONG_TEST_TITLES,
    'outdoor_finishing': OUTDOOR_FINISHING_TEST_TITLES,
    'indoor_chuyen': INDOOR_CHUYEN_TEST_TITLES,
    'indoor_thuong':INDOOR_THUONG_TEST_TITLES,
    'indoor_stone':INDOOR_STONE_TEST_TITLES,
    'indoor_metal':INDOOR_METAL_TEST_TITLES,
    'transit_2c_np': TRANSIT_2C_NP_TEST_TITLES,
    'transit_2c_pallet': TRANSIT_2C_PALLET_TEST_TITLES,
    'transit_RH_np': TRANSIT_RH_NP_TEST_TITLES,
    'transit_RH_pallet': TRANSIT_RH_PALLET_TEST_TITLES,
    # Thêm các nhóm khác nếu cần
}

TEST_TYPE_VI = {
    "ban_us": "BÀN US", "ban_eu": "BÀN EU", "ghe_us": "GHẾ US", "ghe_eu": "GHẾ EU",
    "tu_us": "TỦ US", "tu_eu": "TỦ EU", "giuong": "GIƯỜNG", "guong": "GƯƠNG",
    "indoor_chuyen": "MATERIAL - INDOOR", "indoor_thuong": "MATERIAL - INDOOR",
    "indoor_stone": "MATERIAL - INDOOR", "indoor_metal": "MATERIAL - INDOOR",
    "outdoor_finishing": "MATERIAL - OUTDOOR",
    "testkhac": "TEST KHÁC",
}

# Định nghĩa mapping cho 10 vùng drop test
DROP_LABELS = [
    "Góc 2-3-5", "Cạnh 2-3", "Cạnh 3-5", "Cạnh 2-5",
    "Mặt 1", "Mặt 2", "Mặt 3", "Mặt 4", "Mặt 5", "Mặt 6"
]

DROP_ZONES = [
    "corner_235", "edge_23", "edge_35", "edge_25",
    "face_1", "face_2", "face_3", "face_4", "face_5", "face_6"
]

IMPACT_LABELS = ['Impact 1', 'Impact 2', 'Impact 3', 'Impact 4']
IMPACT_ZONES = ['impact1', 'impact2', 'impact3', 'impact4']
ROT_LABELS = ['Rotation 1', 'Rotation 2', 'Rotation 3', 'Rotation 4']
ROT_ZONES = ['rotation1', 'rotation2', 'rotation3', 'rotation4']

RH_IMPACT_ZONES = [
    ("canh_3_4", "Cạnh 3-4"),
    ("canh_3_6", "Cạnh 3-6"),
    ("canh_4_6", "Cạnh 4-6"),
    ("goc_3_4_6", "Góc 3-4-6"),
    ("goc_2_3_5", "Góc 2-3-5"),
    ("canh_2_3", "Cạnh 2-3"),
    ("canh_1_2", "Cạnh 1-2"),
    ("mat_3_1", "Mặt 3"),
    ("mat_3_2", "Mặt 3 (lần 2)")
]

RH_VIB_ZONES = [
    ("mat_3", "Mặt 3 - 30p"),
    ("mat_4", "Mặt 4 - 15p"),
    ("mat_6", "Mặt 6 - 15p")
]

RH_SECOND_IMPACT_ZONES = [
    ("canh_3_4", "Cạnh 3-4"),
    ("canh_3_6", "Cạnh 3-6"),
    ("canh_1_5", "Cạnh 1-5"),
    ("goc_3_4_6", "Góc 3-4-6"),
    ("goc_1_2_6", "Góc 1-2-6"),
    ("goc_1_4_5", "Góc 1-4-5"),
    ("mat_1", "Mặt 1"),
    ("mat_3", "Mặt 3")
]

for group_dict in TEST_GROUP_TITLES.values():
    group_dict.update(EXTRA_TEST)