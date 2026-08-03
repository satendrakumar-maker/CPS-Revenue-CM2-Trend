from openpyxl import load_workbook
from datetime import datetime
import json
import os

# ═══════════════════════════════════════════════════════════════
# REBUILD SCRIPT — Put this in the same folder as your Excel
# Run: python rebuild.py
# ═══════════════════════════════════════════════════════════════

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)

EXCEL_FILE = "CPS FY 25 26 to Fy 2627.xlsx"
OUTPUT_HTML = "index.html"

print("=" * 55)
print("CPS OEM Dashboard — Rebuild from Excel")
print("=" * 55)
print(f"Reading: {EXCEL_FILE}")

# ── 1. Read Excel ───────────────────────────────────────────
wb = load_workbook(EXCEL_FILE)
ws = wb.active

raw_oems = [
    ("Bajaj (Incl. KTM & TRM)", 3, 6),
    ("Chetak Only", 7, 10),
    ("Ather", 11, 14),
    ("OLA Only CPS", 15, 18),
    ("TVS: All", 19, 22),
    ("Revolt", 23, 26),
    ("Jawa (Incl. Manpower)", 27, 30),
    ("TVS: Only Iqube", 31, 34),
]

data_rows = list(ws.iter_rows(min_row=4, max_row=ws.max_row, values_only=True))

all_months = []
raw_data = {name: {"revenue": [], "retail_count": [], "cm2": [], "cm2_pct": []} for name, _, _ in raw_oems}

for row in data_rows:
    month_val = row[2]
    if month_val is None:
        continue
    if isinstance(month_val, datetime):
        month_label = month_val.strftime("%b-%Y")
    else:
        month_label = str(month_val)
    all_months.append(month_label)

    for name, start, end in raw_oems:
        vals = row[start:end+1]
        raw_data[name]["revenue"].append(float(vals[0]) if vals[0] is not None else 0)
        raw_data[name]["retail_count"].append(float(vals[1]) if vals[1] is not None else 0)
        raw_data[name]["cm2"].append(float(vals[2]) if vals[2] is not None else 0)
        raw_data[name]["cm2_pct"].append(float(vals[3]) if vals[3] is not None else 0)

# ── 2. Computed OEMs ────────────────────────────────────────
def compute_pct(cm2_list, rev_list):
    return [c/r if r != 0 else 0 for c, r in zip(cm2_list, rev_list)]

bajaj_name = "Bajaj (Incl. KTM & TRM)"
chetak_name = "Chetak Only"
tvs_all_name = "TVS: All"
tvs_iqube_name = "TVS: Only Iqube"

bajaj_chetak = {
    "revenue": [raw_data[bajaj_name]["revenue"][i] + raw_data[chetak_name]["revenue"][i] for i in range(len(all_months))],
    "retail_count": [raw_data[bajaj_name]["retail_count"][i] + raw_data[chetak_name]["retail_count"][i] for i in range(len(all_months))],
    "cm2": [raw_data[bajaj_name]["cm2"][i] + raw_data[chetak_name]["cm2"][i] for i in range(len(all_months))],
}
bajaj_chetak["cm2_pct"] = compute_pct(bajaj_chetak["cm2"], bajaj_chetak["revenue"])

tvs_ice = {
    "revenue": [raw_data[tvs_all_name]["revenue"][i] - raw_data[tvs_iqube_name]["revenue"][i] for i in range(len(all_months))],
    "retail_count": [raw_data[tvs_all_name]["retail_count"][i] - raw_data[tvs_iqube_name]["retail_count"][i] for i in range(len(all_months))],
    "cm2": [raw_data[tvs_all_name]["cm2"][i] - raw_data[tvs_iqube_name]["cm2"][i] for i in range(len(all_months))],
}
tvs_ice["cm2_pct"] = compute_pct(tvs_ice["cm2"], tvs_ice["revenue"])

all_oem_data = dict(raw_data)
all_oem_data["Bajaj + Chetak"] = bajaj_chetak
all_oem_data["TVS ICE"] = tvs_ice

total_oem_names = [
    "Bajaj (Incl. KTM & TRM)", "Chetak Only", "Ather",
    "OLA Only CPS", "TVS: All", "Revolt", "Jawa (Incl. Manpower)"
]

# ── 3. FY Totals ────────────────────────────────────────────
fy2526_indices = list(range(0, 12))
fy2627_indices = list(range(12, 15))

def compute_fy_totals(indices):
    total_rev = sum(sum(raw_data[n]["revenue"][i] for i in indices) for n in total_oem_names)
    total_cm2 = sum(sum(raw_data[n]["cm2"][i] for i in indices) for n in total_oem_names)
    total_ret = sum(sum(raw_data[n]["retail_count"][i] for i in indices) for n in total_oem_names)
    avg_pct = total_cm2 / total_rev if total_rev != 0 else 0
    return {"revenue": total_rev, "cm2": total_cm2, "retail_count": total_ret, "cm2_pct": avg_pct}

fy2526_totals = compute_fy_totals(fy2526_indices)
fy2627_totals = compute_fy_totals(fy2627_indices)
overall_totals = compute_fy_totals(fy2526_indices + fy2627_indices)

rev_growth = ((overall_totals["revenue"] - fy2526_totals["revenue"]) / fy2526_totals["revenue"] * 100) if fy2526_totals["revenue"] else 0
ret_growth = ((overall_totals["retail_count"] - fy2526_totals["retail_count"]) / fy2526_totals["retail_count"] * 100) if fy2526_totals["retail_count"] else 0
cm2_growth = ((overall_totals["cm2"] - fy2526_totals["cm2"]) / fy2526_totals["cm2"] * 100) if fy2526_totals["cm2"] else 0

# ── 4. Build data object ────────────────────────────────────
data_obj = {
    "months": all_months,
    "fy_splits": {
        "overall": all_months,
        "fy2526": all_months[:12],
        "fy2627": all_months[12:],
    },
    "fy_totals": {
        "overall": overall_totals,
        "fy2526": fy2526_totals,
        "fy2627": fy2627_totals,
    },
    "growth": {
        "revenue_growth_pct": round(rev_growth, 2),
        "retail_growth_pct": round(ret_growth, 2),
        "cm2_growth_pct": round(cm2_growth, 2),
    },
    "oems": all_oem_data,
    "colors": {
        "Bajaj (Incl. KTM & TRM)": "#E74C3C",
        "Chetak Only": "#3498DB",
        "Ather": "#F39C12",
        "OLA Only CPS": "#1ABC9C",
        "TVS: All": "#9B59B6",
        "Revolt": "#E67E22",
        "Jawa (Incl. Manpower)": "#2ECC71",
        "TVS: Only Iqube": "#34495E",
        "Bajaj + Chetak": "#FF6B6B",
        "TVS ICE": "#00D2D3",
    }
}

data_json_str = json.dumps(data_obj, indent=2)

print(f"   Months: {len(all_months)}")
print(f"   OEMs: {len(all_oem_data)}")
print(f"   FY 25-26 Revenue: ₹{fy2526_totals['revenue']:,.0f}")
print(f"   FY 26-27 Revenue: ₹{fy2627_totals['revenue']:,.0f}")
print(f"   Overall Revenue:  ₹{overall_totals['revenue']:,.0f}")

# ── 5. Read HTML template and inject data ───────────────────
# Use the existing index.html as template, replacing the data block
with open(OUTPUT_HTML, 'r', encoding='utf-8') as f:
    html_content = f.read()

# Find and replace the RAW_DATA line
import re
pattern = r'const RAW_DATA = \{.*?\};'
replacement = f'const RAW_DATA = {data_json_str};'
new_html = re.sub(pattern, replacement, html_content, flags=re.DOTALL)

if new_html == html_content:
    # If regex didn't match, try finding the exact line
    lines = html_content.split('\n')
    for i, line in enumerate(lines):
        if 'const RAW_DATA =' in line:
            # Find the end of the data block (next semicolon after the closing brace)
            start = i
            brace_count = 0
            for j in range(i, len(lines)):
                brace_count += lines[j].count('{') - lines[j].count('}')
                if brace_count == 0 and ';' in lines[j]:
                    end = j
                    break
            lines[start:end+1] = [f'const RAW_DATA = {data_json_str};']
            new_html = '\n'.join(lines)
            break

with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
    f.write(new_html)

print(f"\n✅ Rebuilt: {OUTPUT_HTML}")
print(f"   Size: {len(new_html):,} bytes")
print(f"\n   Next steps:")
print(f"   1. Open {OUTPUT_HTML} in browser to preview")
print(f"   2. Upload {OUTPUT_HTML} to GitHub as index.html")
print(f"   3. Enable GitHub Pages → live dashboard")
