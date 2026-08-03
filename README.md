# CPS OEM Dashboard

Professional dashboard for CPS OEM data — Revenue, CM2, Retail Count, and CM2% trends.

## What's in This Folder

| File | Purpose |
|------|---------|
| `index.html` | **The dashboard** — self-contained, works in any browser |
| `rebuild.py` | **Python script** — rebuilds `index.html` when Excel data changes |
| `CPS FY 25 26 to Fy 2627.xlsx` | **Source Excel file** — edit this, then run rebuild.py |
| `README.md` | This file |

---

## 🚀 Deploy to GitHub Pages

### Step 1: Create GitHub Repo
- Go to [github.com/new](https://github.com/new)
- Name: `cps-oem-dashboard`
- Make it **Public**

### Step 2: Upload Files
Upload **all files in this folder** to your repo:
- `index.html`
- `rebuild.py`
- `CPS FY 25 26 to Fy 2627.xlsx`
- `README.md`

### Step 3: Enable GitHub Pages
- Settings → Pages
- Source: **Deploy from a branch**
- Branch: `main` → folder `/(root)`
- Save → Wait 1 minute

**Live URL:** `https://yourusername.github.io/cps-oem-dashboard/`

---

## 🔄 How to Refresh Data (After Excel Changes)

### Method 1: Run rebuild.py (Recommended)

```bash
# Navigate to this folder
cd "C:\Users\saten\Documents\kimi\workspace\CPS_OEM_Dashboard"

# Run the rebuild script
python rebuild.py
```

This reads the Excel file and regenerates `index.html` with fresh data.

### Method 2: Manual Steps

1. **Edit the Excel file** (`CPS FY 25 26 to Fy 2627.xlsx`) in Excel
2. **Save it** in this same folder
3. **Run:** `python rebuild.py`
4. **Upload the new `index.html`** to GitHub (overwrite old one)
5. **Wait 1 minute** → Dashboard updates automatically

---

## 📊 Dashboard Features

| Feature | Description |
|---------|-------------|
| **Executive Tab** | 6 KPI cards, Revenue/Retail/CM2/CM2% trends, OEM ranking, stacked chart, summary table |
| **OEM-wise Tab** | 10 OEM selector pills, per-OEM charts, growth comparison |
| **Filters** | FY (Overall / 25-26 / 26-27), OEM, Month, Quarter |
| **Dark/Light Mode** | Toggle with 🌙/☀️ button |
| **Chart Download** | Click 💾 on any chart to save as PNG |
| **Export CSV** | Click 📥 Export to download all data |
| **Search Table** | Type in search box to filter OEMs |

---

## ✅ Verified Data

| Period | Revenue | CM2 | Retail |
|--------|---------|-----|--------|
| FY 25-26 | ₹142,532,740 | ₹48,260,715 | 150,111 |
| FY 26-27 | ₹35,485,022 | ₹10,267,791 | 34,506 |
| **Overall** | **₹178,017,762** | **₹58,528,506** | **184,617** |

---

## 🏭 OEMs

| OEM | Type |
|-----|------|
| Bajaj (Incl. KTM & TRM) | Original |
| Chetak Only | Original |
| Ather | Original |
| OLA Only CPS | Original |
| TVS: All | Original |
| Revolt | Original |
| Jawa (Incl. Manpower) | Original |
| TVS: Only Iqube | Subset (inside TVS All) |
| **Bajaj + Chetak** | Computed |
| **TVS ICE** | Computed (TVS All − TVS iQube) |

---

## 💡 Tips

- The dashboard is **one file** — `index.html` contains everything
- Open `index.html` directly in Chrome/Edge to preview locally
- The 🔄 Refresh button in the dashboard just reloads the page
- To get **new data**, you must run `python rebuild.py` after editing Excel
