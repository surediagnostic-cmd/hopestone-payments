---
name: monthly-lis-report
description: Monthly LabsmartLIS export → Excel (4 sheets) + HTML snapshot with targets, charts, and filters
---

You are generating the monthly Sure Diagnostics LIS scorecard. Today is the 1st of the month. Complete the following steps in order.

---

## STEP 1 — Download the current month's LabsmartLIS export

1. Open Chrome and navigate to: https://app.labsmartlis.com/72391405/data_exports/new
2. Log in if prompted (ask the user for credentials if needed).
3. Set the export form:
   - Export type: Cases
   - Date range: 1st to last day of the PREVIOUS calendar month (i.e. the month just completed)
   - All collection centres
4. Submit the form and wait for the download link to appear.
5. Capture the CloudFront download URL from the page using browser JS:
   ```js
   document.querySelector('a[href*="cloudfront"]').href
   ```
6. Fetch the CSV content using browser JS (do NOT use curl/Python urllib — CloudFront blocks sandbox):
   ```js
   const url = '<cloudfront_url>';
   const r = await fetch(url); const txt = await r.text();
   console.log(txt.slice(0, 500)); // verify first few lines
   ```
7. Extract data for processing. Key columns: `Patient`, `Collection centre`, `Investigation`, `Referrer`, `Case Type`, `Fee Paid`, `Fee Due`, `Discount`, `Date`.
   - `Fee Paid + Fee Due` = `Fee Paid` + `Fee Due` for each row.
   - **Name mapping**: Treat `Sure Ilesha, Osun` as the Ijofi branch — keep its name as-is in all outputs.

---

## STEP 2 — Compute aggregations using browser JS

Parse the CSV in the browser and compute all aggregations. Work in batches if needed to avoid output truncation.

### A) Centre Summary (Sheet 1)
For each `Collection centre`:
- `uniquePatients` = count of distinct `Patient` values
- `totalInv` = total number of `Investigation` rows
- `pvtCount` = count of rows where `Referrer` contains "PVT" (case-insensitive)
- `uniqueRef` = count of distinct non-PVT referrer names
- `feeDue` = sum of `Fee Due`
- `discount` = sum of `Discount`
- `feePaidDue` = sum of (`Fee Paid` + `Fee Due`)

### B) Centre-Case Breakdown (Sheet 3)
For each `Collection centre` × `Case Type`:
- `qty` = count of rows
- `feePaidDue` = sum of (`Fee Paid` + `Fee Due`)

### C) Rebate by Referrer (Sheet 2)

For all the centres, group by `Referrer`:
- `feePaid` = sum of `Fee Paid` for that referrer
- Rebate rate = **10%** if ANY investigation in that referrer group matches the regex `\bOS\b` (word boundary, case-insensitive), else **20%**
- `rebate` = feePaid × rate

### D) Monthly Trends
Group by month (`YYYY-MM` extracted from `Date`) and `Collection centre`:
- `patients` = distinct Patient count
- `invCount` = row count
- `referrers` = distinct non-PVT referrer count
- `revenue` = sum of (Fee Paid + Fee Due)

Also group by month and `Case Type`:
- `qty` = row count
- `revenue` = sum of (Fee Paid + Fee Due)

---

## STEP 3 — Build the Excel workbook using openpyxl

Save as: `/sessions/lucid-determined-davinci/mnt/Sure Finance & Business Scorecards/Sure_LIS_<MonthYear>_Dashboard.xlsx`
(e.g. `Sure_LIS_April2026_Dashboard.xlsx` for the April run)

Use these style helpers:
```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def fill(h): return PatternFill('solid', start_color=h, end_color=h)
def fnt(bold=False, sz=10, color='000000', italic=False): return Font(name='Arial', bold=bold, size=sz, color=color, italic=italic)
def aln(h='center', v='center', wrap=False): return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def thin():
    s = Side(style='thin', color='BFBFBF')
    return Border(top=s, bottom=s, left=s, right=s)
def thick():
    s = Side(style='medium', color='404040')
    return Border(top=s, bottom=s, left=s, right=s)

TITLE_FILL = fill('0D2137')
ALT = fill('EBF3FB'); WHT = fill('FFFFFF')
GREEN_FILL = fill('D5F5E3'); AMBER_FILL = fill('FEF9E7'); RED_FILL = fill('FADBD8')
CENTRE_COLORS = {
    'OAUTH Ilesa Centre':       'C55A11',
    'Ilasa Main Centre, Lagos': '2E86AB',
    'Palm Avenue, Lagos':       '16A085',
    'Sure Ilesha, Osun':        'E67E22',
    'Ikeja Lagos Centre':       '8E44AD',
}
```

**Quarterly targets** — use the correct quarter based on the month being reported:
```python
Q_TARGETS = {
    '2026-Q1': {'OAUTH Ilesa Centre':5500000,'Ilasa Main Centre, Lagos':5000000,'Palm Avenue, Lagos':2500000,'Sure Ilesha, Osun':4250000,'Ikeja Lagos Centre':2000000},
    '2026-Q2': {'OAUTH Ilesa Centre':7000000,'Ilasa Main Centre, Lagos':6100000,'Palm Avenue, Lagos':3500000,'Sure Ilesha, Osun':6000000,'Ikeja Lagos Centre':3000000},
    '2026-Q3': {'OAUTH Ilesa Centre':7500000,'Ilasa Main Centre, Lagos':7000000,'Palm Avenue, Lagos':4000000,'Sure Ilesha, Osun':7500000,'Ikeja Lagos Centre':5250000},
    '2026-Q4': {'OAUTH Ilesa Centre':8000000,'Ilasa Main Centre, Lagos':7950000,'Palm Avenue, Lagos':4410000,'Sure Ilesha, Osun':8500000,'Ikeja Lagos Centre':7250000},
}
DAILY_TARGETS = {
    '2026-Q1': {'OAUTH Ilesa Centre':220000,'Ilasa Main Centre, Lagos':200000,'Palm Avenue, Lagos':100000,'Sure Ilesha, Osun':170000,'Ikeja Lagos Centre':80000},
    '2026-Q2': {'OAUTH Ilesa Centre':280000,'Ilasa Main Centre, Lagos':244000,'Palm Avenue, Lagos':140000,'Sure Ilesha, Osun':240000,'Ikeja Lagos Centre':120000},
    '2026-Q3': {'OAUTH Ilesa Centre':300000,'Ilasa Main Centre, Lagos':280000,'Palm Avenue, Lagos':160000,'Sure Ilesha, Osun':300000,'Ikeja Lagos Centre':210000},
    '2026-Q4': {'OAUTH Ilesa Centre':320000,'Ilasa Main Centre, Lagos':318000,'Palm Avenue, Lagos':176400,'Sure Ilesha, Osun':340000,'Ikeja Lagos Centre':290000},
}

def get_quarter(month_str):  # 'YYYY-MM' → 'YYYY-Q1/Q2/Q3/Q4'
    m = int(month_str[5:7])
    q = (m - 1) // 3 + 1
    return f"{month_str[:4]}-Q{q}"

report_month = '<YYYY-MM of reported month>'  # e.g. '2026-04'
quarter_key = get_quarter(report_month)
TARGETS = Q_TARGETS.get(quarter_key, Q_TARGETS['2026-Q1'])
DAILY = DAILY_TARGETS.get(quarter_key, DAILY_TARGETS['2026-Q1'])
```

**Sheet 1 — Centre Summary** (columns A-J):
- Title row spanning A1:J1 with TITLE_FILL, white bold text: "SURE DIAGNOSTICS — {MONTH YEAR}  |  Collection Centre Summary"
- Header row 2: Collection Centre | Unique Patients | Total Investigations | Dr. PVT Cases | Unique Referrers | Fee Due (₦) | Discount (₦) | Fee Paid + Due (₦) | Q Target (₦) | Achievement %
- One data row per centre (alternating ALT/WHT fill)
- Achievement % = feePaidDue / Q_TARGET; number_format = '0.0%'; GREEN if ≥100%, AMBER if ≥70%, RED if <70%
- Grand total row with TITLE_FILL dark blue

**Sheet 2 — Rebate by Referrer** (columns A-E):
- Title row: "SURE DIAGNOSTICS — {MONTH YEAR}  |  Referrer Rebate (Referred Cases)"
- NOTE: Only include the 3 centres that get rebate (Ilasa Main Centre, Palm Avenue, Ikeja Lagos Centre)
- Columns: Collection Centre | Referrer | Fee Paid (₦) | Rebate Rate | Rebate (₦)
- Subtotal row per centre with centre colour fill
- Grand total row

**Sheet 3 — Centre-Case Breakdown** (columns A-D):
- Title row: "SURE DIAGNOSTICS — {MONTH YEAR}  |  Centre-Case Breakdown"
- Columns: Collection Centre | Case Type | Qty | Fee Paid + Due (₦)
- One row per centre × case type combination, sorted by centre then case type
- Subtotal row per centre with centre colour fill
- Grand total row

**Sheet 4 — Target vs Actual** (columns A-G):
- Title row: "SURE DIAGNOSTICS — {MONTH YEAR}  |  Target vs Actual ({Quarter} Targets)"
- Columns: Collection Centre | Q Target (₦) | Actual Revenue (₦) | Variance (₦) | Achievement % | Status | Daily Target (₦)
- Variance = Actual - Target; GREEN fill if positive, RED if negative
- Status: '✓ TARGET MET' if ≥100%, '▲ CLOSE' if ≥70%, '✗ BELOW TARGET' if <70%
- Grand total row

---

## STEP 4 — Build the HTML Snapshot

Save as: `/sessions/lucid-determined-davinci/mnt/Sure Finance & Business Scorecards/Sure_LIS_<MonthYear>_Snapshot.html`

This is a self-contained single HTML file with embedded CSS and JavaScript. It has TWO modes:
1. **Pre-loaded mode** (default): Shows the current month's data embedded as JS constants — no internet required to view.
2. **CSV upload mode**: User can drag-drop or click-upload a new month's CSV to regenerate all tables and charts.

### Structure

**Header**: Dark navy (#0D2137) with "SURE DIAGNOSTICS LIS DASHBOARD — {MONTH YEAR}" title, subtitle showing total rows processed and date generated.

**Upload zone**: Dashed border drag-drop area for CSV upload. Shows "Drop CSV here or click to upload" and a Reset button. When in pre-loaded mode, a green badge shows "✓ Showing {Month Year} pre-loaded data".

**Filter bar** (always visible): 
- Month dropdown (pre-set to current report month in format "Mon YYYY")
- Year dropdown (pre-set to current year)
- Centre multi-select (all 5 centres listed, all selected by default)
- Reset button

### Target vs Actual Cards
Show one card per centre as a horizontal progress bar:
- Card shows: centre name, actual revenue vs target, achievement %, variance, status badge (green/amber/red)
- Progress bar fills proportionally (capped at 100%)
- Colour: green if ≥100%, amber if ≥70%, red if <70%
- Use the correct quarter's targets based on the filtered months

### Tables
**Table 1 — Centre Summary**: Collection Centre | Unique Patients | Total Investigations | Dr. PVT Cases | Unique Referrers | Fee Due (₦) | Discount (₦) | Fee Paid+Due (₦) | Q Target (₦) | Achievement % — with grand total row.

**Table 2 — Centre-Case Breakdown**: Collection Centre | Case Type | Qty | Fee Paid+Due (₦) — grouped by centre with subtotals per centre and grand total.

**Table 3 — Rebate by Referrer**: Collection Centre | Referrer | Fee Paid (₦) | Rebate Rate | Rebate (₦) — only the 3 rebate centres (Ilasa, Palm Ave, Ikeja), with centre subtotals and grand total.

### Charts (using Chart.js v4.4.1 from cdnjs.cloudflare.com)

**Section A — Centre Performance** (4 charts in 2×2 grid):
1. `chartPatients` — Doughnut: Unique Patients by Centre
2. `chartRevenue` — Doughnut: Revenue by Centre  
3. `chartInvestigations` — Bar: Total Investigations by Centre
4. `chartReferrers` — Bar: Unique Referrers by Centre

**Section B — Monthly Trends by Location** (4 stacked bar charts):
5. `chartMthPatients` — Stacked bar: Patients per month, stacked by centre
6. `chartMthInvestigations` — Stacked bar: Investigations per month, stacked by centre
7. `chartMthReferrers` — Stacked bar: Referrers per month, stacked by centre
8. `chartMthRevenueCentre` — Stacked bar: Revenue per month, stacked by centre

**Section C — Monthly Trends by Case Type** (2 stacked bar charts):
9. `chartMthTests` — Stacked bar: Tests per month, stacked by case type
10. `chartMthRevenueCaseType` — Stacked bar: Revenue per month, stacked by case type

Centre colours: OAUTH=#C55A11, Ilasa=#2E86AB, Palm=#16A085, Ilesha=#E67E22, Ikeja=#8E44AD
Case type colours: Lab=#3498DB, USG=#E74C3C, Digital X-ray=#2ECC71, ECG=#F39C12

For stacked bar charts, set `stack: 'stack'` on each dataset and `scales: { x: { stacked: true }, y: { stacked: true } }`.

### Embedded Pre-loaded Data
Embed these JS constants at the top of the `<script>` block with the computed aggregations:

```javascript
const REPORT_MONTH = '<YYYY-MM>';  // e.g. '2026-04'
const REPORT_LABEL = '<Mon YYYY>'; // e.g. 'Apr 2026'
const PRELOADED_TOTAL_ROWS = <N>;

const PRELOADED_CENTRE_AGG = {
  '<Centre Name>': { uniquePatients, totalInv, pvtCount, uniqueRef, feeDue, discount, feePaidDue },
  ...
};

const PRELOADED_BREAKDOWN = {
  '<Centre Name>': { '<Case Type>': { qty, feePaidDue }, ... },
  ...
};

const PRELOADED_REBATE = {
  '<Centre Name>': [ { referrer, feePaid, rebate, rate: '10%'|'20%' }, ... ],
  ...  // only Ilasa, Palm Ave, Ikeja
};

const PRELOADED_MONTHLY = {
  byCentre: { '<YYYY-MM>': { '<Centre>': { patients, invCount, referrers, revenue } } },
  byCaseType: { '<YYYY-MM>': { '<CaseType>': { qty, revenue } } },
};

const Q_TARGETS = {
  '2026-Q1': {'OAUTH Ilesa Centre':5500000,'Ilasa Main Centre, Lagos':5000000,'Palm Avenue, Lagos':2500000,'Sure Ilesha, Osun':4250000,'Ikeja Lagos Centre':2000000},
  '2026-Q2': {'OAUTH Ilesa Centre':7000000,'Ilasa Main Centre, Lagos':6100000,'Palm Avenue, Lagos':3500000,'Sure Ilesha, Osun':6000000,'Ikeja Lagos Centre':3000000},
  '2026-Q3': {'OAUTH Ilesa Centre':7500000,'Ilasa Main Centre, Lagos':7000000,'Palm Avenue, Lagos':4000000,'Sure Ilesha, Osun':7500000,'Ikeja Lagos Centre':5250000},
  '2026-Q4': {'OAUTH Ilesa Centre':8000000,'Ilasa Main Centre, Lagos':7950000,'Palm Avenue, Lagos':4410000,'Sure Ilesha, Osun':8500000,'Ikeja Lagos Centre':7250000},
};
```

### Filter Logic

- `usePreloaded` flag (boolean): true on load, false after CSV upload.
- When `usePreloaded = true`: Month/year filters are locked to pre-loaded month. Centre filter slices `PRELOADED_CENTRE_AGG`, `PRELOADED_BREAKDOWN`, `PRELOADED_REBATE`, and `PRELOADED_MONTHLY` by selected centres only.
- When `usePreloaded = false`: Full pipeline runs on `allRows` array filtered by selected month, year, and centres.

### Quarterly target selection
```javascript
function getQuarter(monthStr) {  // 'YYYY-MM'
  const m = parseInt(monthStr.slice(5));
  const q = Math.ceil(m / 3);
  return monthStr.slice(0, 4) + '-Q' + q;
}
function getTargets(months) {
  if (!months || months.length === 0) return Q_TARGETS['2026-Q1'];
  const mid = months[Math.floor(months.length / 2)];
  return Q_TARGETS[getQuarter(mid)] || Q_TARGETS['2026-Q1'];
}
```

---

## STEP 5 — Verify outputs

1. Open the generated Excel file and confirm all 4 sheets exist with data.
2. Open the HTML snapshot in the browser — check that:
   - Pre-loaded data renders automatically without any uploads
   - Filter bar shows the correct month pre-selected
   - Target vs Actual progress cards display with correct colours
   - All 10 charts render (no blank charts)
   - Centre-Case Breakdown table shows subtotals per centre
   - Rebate table shows only the 3 rebate centres
3. Test CSV upload with the same export to verify the "upload mode" also works.

---

## IMPORTANT RULES

- **Rebate rate**: 10% if any investigation for that referrer contains `\bOS\b` (regex, word boundary); otherwise 20% of Fee Paid.
- **PVT classification**: Referrer contains "PVT" (case-insensitive) = private patient.
- **Corporate/HMO classification**: Referrer contains "CORPORATE" or "HMO" = corporate patient.
- **Referral classification**: All other referrer values = referral.
- **Sure Ilesha, Osun** = Ijofi branch — use this name as-is in all outputs.
- Save both files to: `/sessions/lucid-determined-davinci/mnt/Sure Finance & Business Scorecards/`
- After saving, present download links for both files to the user.