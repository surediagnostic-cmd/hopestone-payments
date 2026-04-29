---
name: sure-monthly-dashboard
description: "Run the Sure Finance & Business monthly performance dashboard pipeline. Use this skill whenever the user wants to generate a monthly dashboard, process LabSmartLIS data for a full month, produce monthly performance reports, view branch performance for a specific month, or filter dashboard data by month, year, and branch. Also triggers when the user mentions monthly totals, monthly summaries, monthly targets, generating the monthly Excel or HTML report, refreshing the monthly scorecard, viewing revenue by case type, viewing tests completed by location, viewing patients by location, or active referrers by location. Use this skill for both monthly breakdowns and yearly comparisons across all months."
---

# Sure Finance Monthly Dashboard Pipeline

Generates the **monthly** performance dashboard from a LabSmartLIS cases CSV export. The only output is:
- **Multi-period HTML** (`Sure_Monthly_Dashboard_Multi.html`) — Monthly (Jan–Dec) and Yearly (2026, 2027…) toggle, Year/Month/Branch filters, dark/light mode, P&L section

> The Excel workbook and single-month HTML are no longer generated. Only the multi-period HTML is produced each run.

## Charts produced — Multi-period HTML

### Filters & views

Controls: **Month** · **Year** · **Branch** · **[Monthly/Yearly toggle]** · **Dark/Light mode button**

| Filter / Toggle | Behaviour |
|-----------------|-----------|
| Month = All Months | X-axis = Jan–Dec for the selected year; KPI shows monthly average |
| Month = specific (e.g. Feb) | X-axis = that month only; KPI shows that month's actuals vs monthly target |
| Yearly toggle (active) | X-axis = 2026, 2027…; Month filter hidden; KPI = all-time totals |
| Year dropdown | Selects which year's data to show (Monthly mode) |
| Branch dropdown | Filters all charts and KPI to a single branch or All Branches |

All charts respond to every filter change simultaneously.

### Performance charts (6)

All stacked bar charts display a **column total label** above each bar (format: `1.23k` / `1.23M`). The label is drawn by a registered Chart.js plugin (`stackedTotals`) — this runs on every render including theme switches.

| Chart | X-axis | Notes |
|-------|--------|-------|
| Revenue by Location | Months (Jan–Dec) or years | **Stacked bar + dashed orange target line**; stacked totals shown |
| Tests Completed by Location | Months or years | Stacked bar by branch; stacked totals shown |
| No. of Patients by Location | Months or years | Stacked bar by branch; stacked totals shown |
| Active Referrers by Location | Months or years | Unique referrer count per branch; stacked totals shown |
| Revenue by Case Type | Case types (Lab, USG…) | Stacked bar — each dataset = a branch; stacked totals shown |
| Top Agents | Agent names (horizontal bar) | Top 10 by revenue; stacked totals shown at right end of bar |

### P&L charts (3) — shown only when `cost_history.json` has data

| Chart | Type | Notes |
|-------|------|-------|
| P&L — Net Profit by Branch | Line chart (one line per branch) | Revenue − Direct Costs − OpEx per period |
| P&L — Cost Breakdown by Branch | Grouped bar | Revenue, Direct Costs, Overhead/OpEx, Net Profit side by side per branch |
| P&L — Cost Breakdown / Net Profit (Branches Stacked) | Stacked bar | **Dynamic X-axis** — see below; stacked totals shown |

**P&L — Cost Breakdown (Branches Stacked) — dynamic X-axis logic:**

| View mode | X-axis | Y value | Chart title |
|-----------|--------|---------|-------------|
| Specific month selected | Revenue · Direct Costs · Overhead/OpEx · Net Profit | Aggregated value per category, branches stacked | P&L — Cost Breakdown (Branches Stacked) |
| All Months (monthly mode) | Jan · Feb · … · Dec | Net Profit per month, branches stacked | P&L — Net Profit by Month [YYYY] (Branches Stacked) |
| Yearly toggle | 2026 · 2027 · … | Net Profit per year, branches stacked | P&L — Net Profit by Year (Branches Stacked) |

Both the grouped and stacked P&L charts carry a reconciliation note: revenue uses LabSmartLIS totals (at case creation); costs come from management P&L reports. Actual net profit may differ by ~5–10% if P&L revenue has been adjusted.

### KPI card

Single card (not one per branch), updates with every filter change.
- *All Branches*: aggregates all centres
- *Specific branch*: shows that branch's numbers only
- When `cost_history.json` has data for the period: also shows Direct Costs, Gross Profit, Overhead/OpEx, Net Profit tiles (colour-coded green/red)

**Target logic (important)**:
- All Months KPI → shows **average monthly** figures vs the **monthly target** (×1)
- Specific month KPI → shows **that month's actuals** vs the **monthly target** (×1)
- Yearly view KPI → shows **all-time totals** vs the **annual target** (monthly target × 12 × n_years)

**Revenue target line on Revenue chart**:
- Monthly view: dashed orange line at the total monthly target for each month (branch-filtered or all-branches sum)
- Yearly view: flat dashed orange line at the annual target

**Active Referrers = unique count**: Computed from distinct referrer names in the `referrers` list (not a sum of daily unique_referrers fields). Each referrer name is only counted once per period.

**Dark/light mode**: Auto-detects OS preference on first load. Manual toggle button (top-right of header) persists the choice in `localStorage`. All chart grid lines and tick labels update on theme change.

---

## Key constraints

- **Pipeline script**: `update_monthly_dashboard.py` in the session working directory
- **Permanent outputs**: the "Sure Finance & Business Scorecards" workspace folder
- **History file**: `monthly_dashboard_history.json` in `Documents/Sure Finance Dashboards/` — persists across sessions, do NOT delete
- **Cost history file**: `cost_history.json` in `Documents/Sure Finance Dashboards/` — keyed as `"YYYY-MM"`, each entry is `{"Branch name": {"direct_costs": N, "opex": N}}`. Updated manually from management P&L reports. Do NOT delete.
- **No browser required**: CSV upload → Python handles everything. No CDN fetching, no JS chunking.
- **Centre mapping**: "Sure Ilesha, Osun" → "Ijofi" branch targets (handled automatically by script)
- **Timezone**: WAT (UTC+1)
- **Branch colours**: Ilasa=gold (#F4B942), OAUTH Ilesa=green (#4CAF50), Palm Avenue=purple (#9C27B0), Ijofi=orange (#FF7043), Ikeja=blue (#2196F3)

---

## Step 0 — Clarify the target month (ask if not specified)

Ask the user:
1. **Which month and year?** (e.g. "March 2026")
2. **Do you have the CSV ready?** If not, guide them through the export (Step 1). If they already uploaded it, go straight to Step 2.

---

## Step 1 — Export CSV from LabSmartLIS

If the user has not yet uploaded a CSV, ask them to:

1. Log in to **https://app.labsmartlis.com/72391405**
2. Go to **Business → Data Export**
3. Set Entity = **Case**, Date range = **1st of the month 00:00 → last day 23:59** (e.g. 2026-03-01 to 2026-03-31)
4. Click **Run export**, wait for it to complete, then **download the CSV**
5. **Upload the CSV file directly into this conversation**

> The pipeline accepts the raw CSV directly — no aggregation or browser steps needed.

---

## Step 2 — Run the pipeline

Once the CSV is uploaded (path will be under `/sessions/.../mnt/uploads/`), run to a temp directory first to avoid macOS FUSE file-lock errors (the existing dashboard may be open in a browser):

```bash
pip install openpyxl --break-system-packages -q

mkdir -p /tmp/sure_dash

python3 /sessions/<session-id>/update_monthly_dashboard.py \
  --csv-file /sessions/<session-id>/mnt/uploads/<filename>.csv \
  --month 3 --year 2026 \
  --workspace /tmp/sure_dash
```

Then copy the output to the Scorecards folder. If the old file is locked, use a new name:

```bash
# If Sure_Monthly_Dashboard_Multi.html is not locked:
cp /tmp/sure_dash/Sure_Monthly_Dashboard_Multi.html \
   "/sessions/<session-id>/mnt/Sure Finance & Business Scorecards/Sure_Monthly_Dashboard_Multi.html"

# If the file is locked (OSError: Resource deadlock avoided), use a versioned name:
cp /tmp/sure_dash/Sure_Monthly_Dashboard_Multi.html \
   "/sessions/<session-id>/mnt/Sure Finance & Business Scorecards/Sure_PnL_Dashboard_Multi_v2.html"
```

Replace `<session-id>` with the actual session directory (use `ls /sessions/` to find it).

**The `--csv-file` flag handles everything**: reads the raw CSV, aggregates by date/centre/referrer/agent/case-type, then runs the full pipeline in one step.

**Alternative flags** (for pre-aggregated data from a previous run):
- `--data-file /tmp/monthly_data.json` — pass a pre-aggregated JSON file
- `--data '{"daily":[...],...}'` — pass JSON inline

---

## Step 3 — Outputs

| File | Location | Description |
|------|----------|-------------|
| `Sure_Monthly_Dashboard_Multi.html` | Sure Finance & Business Scorecards folder | **The only output** — multi-period dashboard with dark/light mode, target lines, P&L section |
| `monthly_dashboard_history.json` | Documents/Sure Finance Dashboards/ | Cumulative LabSmartLIS history — never delete, grows each month |
| `cost_history.json` | Documents/Sure Finance Dashboards/ | P&L cost data (direct costs + opex per branch per month) — never delete, updated manually |

Provide the user with a clickable `computer://` link to the HTML file.

---

## Step 4 — Refresh the Cowork sidebar artifact

After Step 3 completes, always refresh the **`sure-finance-monthly-dashboard`** Cowork artifact so the sidebar stays in sync without any manual action.

Run this Python patch script in the session:

```python
import json, re, subprocess, sys

# ── 1. Read the freshly generated HTML ─────────────────────────────────────
with open('/tmp/sure_dash/Sure_Monthly_Dashboard_Multi.html', encoding='utf-8') as f:
    html = f.read()

# ── 2. Swap CDN to artifact-approved version (jsdelivr + no SRI issues) ────
html = html.replace(
    'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js',
    'https://cdn.jsdelivr.net/npm/chart.js@4.5.0/dist/chart.umd.js'
)

# ── 2b. Add --brand CSS variable so blue text becomes white in dark mode ───
html = html.replace(
    '--shadow:rgba(0,0,0,.07);--shadow2:rgba(0,0,0,.05);\n}',
    '--shadow:rgba(0,0,0,.07);--shadow2:rgba(0,0,0,.05);--brand:#1F3864;\n}'
)
html = html.replace(
    '--shadow:rgba(0,0,0,.3);--shadow2:rgba(0,0,0,.2);\n}\n@media',
    '--shadow:rgba(0,0,0,.3);--shadow2:rgba(0,0,0,.2);--brand:#e6edf3;\n}\n@media'
)
html = html.replace(
    '--shadow:rgba(0,0,0,.3);--shadow2:rgba(0,0,0,.2);\n}}',
    '--shadow:rgba(0,0,0,.3);--shadow2:rgba(0,0,0,.2);--brand:#e6edf3;\n}}'
)
html = html.replace('color:#1F3864}', 'color:var(--brand)}')
html = html.replace('color:#1F3864;', 'color:var(--brand);')

# ── 3. Strip weekly sub-objects from HISTORY to keep artifact size small ───
hist_match = re.search(r'const HISTORY\s*=\s*(\{.*?\});', html, re.DOTALL)
if hist_match:
    history = json.loads(hist_match.group(1))
    KEEP = ('investigations','cases','patients','revenue','paid','due','unique_referrers')
    stripped = {}
    for k, v in history.items():
        centres = {
            cn: {fk: cd[fk] for fk in KEEP if fk in cd}
            for cn, cd in v.get('centres', {}).items()
        }
        stripped[k] = {
            'month': v['month'], 'year': v['year'],
            'centres': centres,
            'case_types': v.get('case_types', []),
            'agents':     v.get('agents', []),
            'referrers':  v.get('referrers', []),
        }
    html = (html[:hist_match.start()]
            + 'const HISTORY      = '
            + json.dumps(stripped, separators=(',',':'))
            + ';'
            + html[hist_match.end():])

# ── 4. Remove localStorage (not supported inside Cowork artifacts) ─────────
html = html.replace("localStorage.getItem('theme')", "null")
html = html.replace("localStorage.setItem('theme',next);\n  ", "")

# ── 5. Fix KPI ordering — ensure KPI runs BEFORE chart rendering ───────────
#    (Protects against chart errors silently swallowing KPI updates)
def move_kpi_first(func_body, kpi_comment, chart_comment, kpi_call_prefix):
    try:
        cs = func_body.index(chart_comment)
        ks = func_body.index(kpi_comment)
        kc = func_body.index(kpi_call_prefix, ks)
        ke = func_body.index('\n', kc + 1) + 1
        if ks > cs:   # KPI is after charts — swap needed
            return func_body[:cs] + func_body[ks:ke] + '\n' + func_body[cs:ks] + func_body[ke:]
    except ValueError:
        pass
    return func_body

rm_s = html.index('function renderMonthly(')
ry_s = html.index('function renderYearly(')
rend = html.index('function render', ry_s + 1)   # next function after renderYearly

rm_body = html[rm_s:ry_s]
ry_body = html[ry_s:rend]

rm_body = move_kpi_first(
    rm_body,
    '\n  // --- KPI card ---\n',
    '\n  // Revenue chart with target line overlay\n',
    'renderKPICard('
)
ry_body = move_kpi_first(
    ry_body,
    '\n  // --- KPI card ---\n',
    '\n  // Revenue chart\n',
    'renderKPICard('
)

html = html[:rm_s] + rm_body + ry_body + html[rend:]

# ── 6. Save patched HTML to session temp path ──────────────────────────────
out_path = '/sessions/vigilant-modest-ritchie/artifact_refresh.html'
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'Patched HTML saved to {out_path}  ({len(html):,} chars)')
```

After the script runs successfully, call `mcp__cowork__update_artifact` with:

```
id:             sure-finance-monthly-dashboard
html:           <full contents of /sessions/vigilant-modest-ritchie/artifact_refresh.html>
update_summary: "Added [Month Year] data — [N] months of history embedded"
```

Read `artifact_refresh.html` with the `Read` tool and pass its content as the `html` parameter. Do **not** summarise or truncate — pass the full string.

> **If `update_artifact` fails** with a size error, the patched file is too large. Strip the `agents` list from all history entries (keep only top-20 by revenue) and retry.

---

## Targets 2026 (monthly target per quarter)

Targets escalate each quarter. The pipeline selects the correct target based on the month being processed.

| Centre | Q1 (Jan–Mar) | Q2 (Apr–Jun) | Q3 (Jul–Sep) | Q4 (Oct–Dec) |
|--------|-------------|-------------|-------------|-------------|
| Ilasa Main Centre, Lagos | NGN 5,000,000 | NGN 6,100,000 | NGN 7,000,000 | NGN 7,950,000 |
| OAUTH Ilesa Centre | NGN 5,500,000 | NGN 7,000,000 | NGN 7,500,000 | NGN 8,000,000 |
| Palm Avenue, Lagos | NGN 2,500,000 | NGN 3,500,000 | NGN 4,000,000 | NGN 4,410,000 |
| Sure Ilesha, Osun (= Ijofi) | NGN 4,250,000 | NGN 6,000,000 | NGN 7,500,000 | NGN 8,500,000 |
| Ikeja Lagos Centre | NGN 2,000,000 | NGN 3,000,000 | NGN 5,250,000 | NGN 7,250,000 |

In the script these are stored as `TARGETS_2026 = {"Centre": {1: Q1_target, 2: Q2_target, 3: Q3_target, 4: Q4_target}}` and retrieved via `get_month_target(centre, month)` which maps the month to its quarter automatically.

---

## Common issues

| Problem | Fix |
|---------|-----|
| Script not found | Use `ls /sessions/*/update_monthly_dashboard.py` to locate it |
| Wrong month in output | Confirm the CSV date range matches the requested month |
| Missing case types | Ensure CSV has a "Case Type" column (some older exports may not) |
| Palm Avenue rows have trailing space | Script normalises via `strip()` — no manual fix needed |
| UnicodeEncodeError | No emoji in JS string literals inside Python f-strings/heredocs |
| Branch filter not matching | Script normalises to lowercase + strip + handles Ilesha→Ijofi alias |
| History wrong format | Old format has `{"runs":[...]}` — reset to `{}` via Python, not file deletion |
| `unique_referrers` KeyError | Ensure `process_data()` passes the field through in `daily_out` |
| `OSError: [Errno 35] Resource deadlock avoided` | Dashboard is open in a browser, locking the mnt file. Use `--workspace /tmp/sure_dash` and copy with a new name (e.g. `_v2.html`) |
| P&L charts not appearing | Check `cost_history.json` exists and has an entry for the selected month key (e.g. `"2026-03"`) |
| Stacked bar totals appear mid-bar instead of at the top | Do NOT use `yScale.getPixelForValue(total)` for label placement — use the actual `bar.y` element position (min canvas-y across positive segments). The scale-value approach places labels at the net-total level, which is wrong for mixed positive/negative stacks and can land mid-bar on pure-positive charts due to y-axis headroom |
| Artifact update fails with size error | Strip `agents` list to top-20 by revenue per history entry and retry Step 4 |

---

## Pipeline Script Specification

Only needed when `update_monthly_dashboard.py` is absent. Recreate it with these requirements:

### Arguments
`--csv-file` (raw CSV path) | `--data` (JSON string) | `--data-file` (JSON path) | `--month` (int, required) | `--year` (int, required) | `--branch` (optional) | `--workspace` (optional, override output directory — default is the Scorecards mnt folder)

When `--csv-file` is supplied, run `aggregate_csv(path)` first to produce the `raw` dict.

### Data model (after aggregation)
```
daily:      [{Date, Collection centre, investigations, cases, patients, revenue, paid, due, unique_referrers}]
referrers:  [{Collection centre, Referrer, cases, investigations, revenue}]
agents:     [{Agent, cases, investigations, patients, revenue, paid}]
case_types: [{Collection centre, Case Type, cases, investigations, revenue}]
```

### `aggregate_csv(csv_path)` — CSV aggregation logic
- Read with `csv.DictReader`, encoding `utf-8-sig`
- Skip rows where `Canceled == "true"`
- Each row = 1 investigation (count rows, not the "Investigations" text field)
- Daily key = `(Date, Collection centre)` — track unique case IDs, unique patient names, unique referrer names as sets
- Referrers key = `(Collection centre, Referrer)`
- Agents key = `Agent` (skip empty agents)
- Case types key = `(Collection centre, Case Type)`
- Revenue fields: `Total Fee`, `Fee Paid`, `Fee Due` → convert to float, default 0

### Core pipeline functions

**`process_data(data, month, year, branch_filter=None)`**
- Parse `Date` as `"%Y-%m-%d"` to filter by month/year
- Map "Sure Ilesha, Osun" → "Ijofi" via `CENTRE_ALIAS`
- Weekly buckets: Week 1=days 1–7, Week 2=8–14, Week 3=15–21, Week 4=22+
- Unique referrers per centre = distinct referrer names in `refs_out` (not sum of daily fields)

**`load_cost_history()` / `save_cost_data(month, year, cost_data, cost_history)`**
- Reads/writes `cost_history.json` in `Documents/Sure Finance Dashboards/`
- Key format: `"YYYY-MM"` (e.g. `"2026-03"`)
- Structure: `{"YYYY-MM": {"Branch name": {"direct_costs": N, "opex": N}}}`

**`build_excel(summary, path, cost_data=None)`** — retained in script but NOT called by `main()`. Generates a 6-sheet workbook including a P&L Summary sheet. Only rebuild/call if an Excel output is explicitly requested.

**`build_html_single(summary, path, cost_data=None)`** — retained in script but NOT called by `main()`. Generates a single-month HTML. Only rebuild/call if an individual-month HTML is explicitly requested.

**`build_html_multi(history, path, cost_history=None)`** — the only function called by `main()`. Generates the multi-period HTML with:
- Monthly/Yearly toggle, Year/Month/Branch filter controls
- Revenue chart uses `mkRevChart` / `mkRevChartYearly` helpers that add a dashed orange target line dataset on top of the stacked bar datasets
- Dark/light mode: CSS custom properties (`:root` light variables, `[data-theme="dark"]` dark variables, `@media(prefers-color-scheme:dark)` auto-detect); `toggleTheme()` JS function with `localStorage` persistence; theme applied on `init()` before first render; all chart tick/grid colours updated on theme change via `getEffectiveTheme()`
- **`stackedTotalsPlugin`**: custom Chart.js plugin registered globally via `Chart.register()`. Draws the net column total above each stacked bar (or to the right for horizontal bars). Key implementation details:
  - Skips datasets with `type:'line'` (e.g. the revenue target line) from both the total sum and the anchor position search
  - **Anchor position uses actual rendered bar element coordinates**, NOT `yScale.getPixelForValue(total)` — this ensures the label always sits at the visual top of the bar regardless of y-axis headroom or Chart.js internal scaling
  - Vertical bars: scans all bar dataset elements per column and tracks the minimum `bar.y` (topmost canvas y-coordinate) among **positive-value segments only** — this correctly handles mixed positive/negative stacks (e.g. Net Profit) by anchoring at the top of the positive portion, not the net total level
  - Horizontal bars: scans all bar dataset elements per row and tracks the maximum `bar.x` (rightmost canvas x-coordinate); label placed at `maxX + 4` with `textBaseline: 'middle'`
  - Enabled per-chart via `plugins.stackedTotals.enabled: true` in chart options
  - Applied to: all 6 performance charts and all P&L stacked charts; NOT applied to the grouped P&L cost breakdown chart (`cPnLCost`)
- **`fmtCompact(v)`**: formats numbers as `1.23k` (≥1,000) or `1.23M` (≥1,000,000), with sign for negatives. Used for stacked bar total labels.
- P&L section (hidden when no cost data): Net Profit line chart + Cost Breakdown grouped bar + Cost Breakdown stacked chart (dynamic — see X-axis logic table above); first two charts carry reconciliation notes
- P&L stacked chart (`cPnLStack`) title updates dynamically via `document.getElementById('pnlStackTitle').textContent`
- KPI card extended with Direct Costs, Gross Profit, Overhead/OpEx, Net Profit tiles when cost data is available

**`load_history()` / `save_to_history(history, summary)`**
- Keyed as `"YYYY-MM"` (e.g. `"2026-03"`)
- Detect old `{"runs":[...]}` format and reset to `{}`

### Key coding constraints
- Use `openpyxl` for Excel (no xlsxwriter)
- Chart.js from `https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js`
- No emoji in JS string literals inside Python f-strings/heredocs
- WORKSPACE = `Path("/sessions/<session-id>/mnt/Sure Finance & Business Scorecards")`
- COST_HISTORY_FILE = `Path("/sessions/<session-id>/mnt/Documents/Sure Finance Dashboards/cost_history.json")`
