# CPS OEM Dashboard

A professional, self-contained HTML dashboard for CPS OEM data visualization.

## What's Included

| File | Purpose |
|------|---------|
| `index.html` | **The dashboard** — contains everything (data, styles, charts, logic). Just open in browser or deploy to GitHub Pages. |
| `data.json` | Data export (for reference / debugging). |
| `CPS FY 25 26 to Fy 2627.xlsx` | Source Excel file. |

## Deploy to GitHub Pages (3 Steps)

1. **Upload `index.html`** to your GitHub repository root
2. Go to **Settings → Pages** → Select **Deploy from a branch** → Choose `main` → Folder `/(root)`
3. Wait 1 minute → Your dashboard is live at:
   ```
   https://yourusername.github.io/your-repo-name/
   ```

That's it. Only **one file** needed.

## Dashboard Features

- **Executive Tab:** 6 KPI cards, Revenue/Retail/CM2/CM2% trends, OEM ranking, stacked charts, summary table
- **OEM-wise Tab:** 10 OEM selector pills, per-OEM 4 charts, growth comparison
- **Filters:** FY (Overall / 25-26 / 26-27), OEM, Month, Quarter
- **Extras:** Dark/Light mode, chart download PNG, export CSV, search table, responsive

## Refresh Data After Excel Changes

```bash
# Step 1: Update your Excel file (same filename)
# Step 2: Run the rebuild script
cd "C:\Users\saten\Documents\kimi\workspace"
python build_standalone.py

# Step 3: Re-upload the new index.html to GitHub
```

Or manually:
1. Edit the Excel file
2. Run `python build_standalone.py` in the workspace folder
3. The new `index.html` is regenerated with fresh data embedded
4. Upload `CPS_OEM_Dashboard/index.html` to GitHub (overwrite)

## Verified Data

| Period | Revenue | CM2 | Retail |
|--------|---------|-----|--------|
| FY 25-26 | ₹142,532,740 | ₹48,260,715 | 150,111 |
| FY 26-27 | ₹35,485,022 | ₹10,267,791 | 34,506 |
| **Overall** | **₹178,017,762** | **₹58,528,506** | **184,617** |

## OEMs

**Original 7:** Bajaj, Chetak, Ather, OLA, TVS (All), Revolt, Jawa
**Computed 2:** Bajaj + Chetak, TVS ICE (= TVS All − TVS iQube)
**Subset 1:** TVS iQube (already included in TVS All totals)

## Local Preview

Simply double-click `index.html` in File Explorer. It opens in Chrome/Edge and works instantly — no server needed.
