#!/bin/bash
# Installs the sure-weekly-dashboard skill into your Claude skills directory.
# Double-click this file in Finder to run it (macOS will open a Terminal window).

set -e

echo "=== Installing sure-weekly-dashboard skill ==="

# Find the Claude skills directory by locating the existing monthly-sure-scorecard skill
SKILLS_DIR=$(find /Users -path "*/monthly-sure-scorecard/SKILL.md" 2>/dev/null | head -1 | xargs dirname 2>/dev/null | xargs dirname 2>/dev/null)

if [ -z "$SKILLS_DIR" ]; then
  # Fallback: standard Cowork location
  SKILLS_DIR="$HOME/.claude/skills"
fi

echo "Skills directory: $SKILLS_DIR"

SKILL_DIR="$SKILLS_DIR/sure-weekly-dashboard"
mkdir -p "$SKILL_DIR"
echo "Created: $SKILL_DIR"

# Write SKILL.md
cat > "$SKILL_DIR/SKILL.md" << 'SKILLEOF'
---
name: sure-weekly-dashboard
description: >
  Run the Sure Finance & Business weekly performance dashboard pipeline. Use this skill whenever the user
  wants to generate the weekly dashboard, process LabSmartLIS data, update performance reports, run the
  weekly pipeline, or refresh the Sure Finance dashboard. Also triggers when the user mentions exporting
  cases from LabSmartLIS, updating dashboard history, or generating the Excel/HTML performance reports.
---

# Sure Finance Weekly Dashboard Pipeline

This skill generates the weekly performance dashboard for Sure Finance & Business from LabSmartLIS data.
It produces an Excel workbook (6 sheets), a single-week HTML dashboard, and a multi-period HTML dashboard
with year/month/week filters.

## Key facts

- **Pipeline script**: `update_dashboard.py` in the session working directory (copy kept in workspace)
- **Permanent outputs**: the user's "Sure Finance & Business Scorecards" workspace folder
- **History file**: `dashboard_history.json` in that same folder — persists across sessions
- **CDN constraint**: LabSmartLIS downloads from `d33hv1py9n6skn.cloudfront.net`, which is blocked
  from the Python sandbox. All data must be fetched via JavaScript inside the Chrome tab.
- **Centre mapping**: "Sure Ilesha, Osun" in LabSmartLIS = Ijofi branch targets in the script.
- **Timezone**: WAT (UTC+1, Africa/Lagos). All scheduling is WAT.
- **Scheduled run**: Cron `59 23 * * 0` = 11:59 PM WAT every Sunday.

---

## Step 1 — Export the CSV from LabSmartLIS

Open Chrome at **https://ng.labsmart.net**.

Navigate: **Business → Data Export**

Filters:
- Entity: Case
- Date range: Monday 00:00 – Sunday 23:59 for the target week
- Click Export / Download

The download link will be a CDN URL (`d33hv1py9n6skn.cloudfront.net/...`).
**Do NOT attempt to fetch it from Python/sandbox** — the CDN is blocked. Use the Chrome tab instead.

---

## Step 2 — Fetch and aggregate data in the browser

Use `mcp__Claude_in_Chrome__javascript_tool` in the active LabSmartLIS tab.

### Get the download URL first

```javascript
// Look for the download link on the Data Export page
Array.from(document.querySelectorAll('a[href*="cloudfront"]')).map(a=>a.href)
```

### Run the aggregation script

```javascript
(async () => {
  const url = 'PASTE_CDN_URL_HERE';
  const resp = await fetch(url);
  const text = await resp.text();
  const lines = text.split('\n');
  const headers = lines[0].split(',').map(h=>h.trim().replace(/^"|"$/g,''));
  function col(row,name){const i=headers.indexOf(name);return i>=0?(row[i]||'').replace(/^"|"$/g,'').trim():'';}
  const rows=[];
  for(let i=1;i<lines.length;i++){const r=lines[i].match(/(".*?"|[^,]+)(?=,|$)/g);if(r&&r.length>=5)rows.push(r);}
  const daily={},referrers={},agents={},caseTypes={};
  for(const r of rows){
    const cancelled=col(r,'Canceled').toLowerCase();
    if(cancelled==='true'||cancelled==='1')continue;
    const date=col(r,'Date').split(' ')[0];
    const centre=col(r,'Collection centre');
    const key=date+'|'+centre;
    const rev=parseFloat(col(r,'Total Fee'))||0;
    const paid=parseFloat(col(r,'Fee Paid'))||0;
    const due=parseFloat(col(r,'Fee Due'))||0;
    const inv=col(r,'Investigations');
    const invCount=inv?inv.split(',').filter(x=>x.trim()).length||1:0;
    const patient=col(r,'Reg. no.');
    const referrer=col(r,'Referrer');
    const agent=col(r,'Agent')||'Walk-in / Unassigned';
    const caseType=col(r,'Case Type');
    if(!daily[key])daily[key]={date,centre,inv:0,cases:0,patients:new Set(),rev:0,paid:0,due:0,refs:new Set()};
    daily[key].inv+=invCount;daily[key].cases++;daily[key].patients.add(patient);
    daily[key].rev+=rev;daily[key].paid+=paid;daily[key].due+=due;daily[key].refs.add(referrer);
    const rk=centre+'|'+referrer;
    if(!referrers[rk])referrers[rk]={centre,referrer,cases:0,inv:0,rev:0};
    referrers[rk].cases++;referrers[rk].inv+=invCount;referrers[rk].rev+=rev;
    if(!agents[agent])agents[agent]={agent,cases:0,inv:0,patients:new Set(),rev:0,paid:0};
    agents[agent].cases++;agents[agent].inv+=invCount;agents[agent].patients.add(patient);
    agents[agent].rev+=rev;agents[agent].paid+=paid;
    const ctk=date+'|'+centre+'|'+caseType;
    if(!caseTypes[ctk])caseTypes[ctk]={date,centre,caseType,cases:0,inv:0};
    caseTypes[ctk].cases++;caseTypes[ctk].inv+=invCount;
  }
  const dailyArr=Object.values(daily).map(d=>({'Date':d.date,'Collection centre':d.centre,investigations:d.inv,cases:d.cases,patients:d.patients.size,revenue:Math.round(d.rev),paid:Math.round(d.paid),due:Math.round(d.due),unique_referrers:d.refs.size}));
  const refArr=Object.values(referrers).map(r=>({'Collection centre':r.centre,Referrer:r.referrer,cases:r.cases,investigations:r.inv,revenue:Math.round(r.rev)})).sort((a,b)=>b.cases-a.cases);
  const agentArr=Object.values(agents).map(a=>({Agent:a.agent,cases:a.cases,investigations:a.inv,patients:a.patients.size,revenue:Math.round(a.rev),paid:Math.round(a.paid)})).sort((a,b)=>b.investigations-a.investigations);
  const ctArr=Object.values(caseTypes);
  const result={daily:dailyArr,referrers:refArr,agents:agentArr,caseTypes:ctArr};
  const str=JSON.stringify(result);
  const CHUNK=200000;const nChunks=Math.ceil(str.length/CHUNK);
  sessionStorage.setItem('dash_nchunks',nChunks);
  for(let i=0;i<nChunks;i++)sessionStorage.setItem('dash_chunk_'+i,str.slice(i*CHUNK,(i+1)*CHUNK));
  return 'Stored '+nChunks+' chunks, '+dailyArr.length+' daily rows, '+refArr.length+' referrers, '+agentArr.length+' agents';
})();
```

### Retrieve chunks

```javascript
// Step 1 — get chunk count
parseInt(sessionStorage.getItem('dash_nchunks')||'0')
// Step 2 — get each chunk (repeat for chunk_1, chunk_2, ...)
sessionStorage.getItem('dash_chunk_0')
```

Concatenate all chunks in Python, then `json.loads()` to get `{daily, referrers, agents, caseTypes}`.

---

## Step 3 — Construct a CSV and run the pipeline

The pipeline script can auto-detect a CSV in the working directory. If the raw CSV was downloaded
into the session folder, just run:

```bash
python3 update_dashboard.py /path/to/cases.csv
```

If only the browser-aggregated JSON is available, reconstruct a minimal CSV from it:

```python
import pandas as pd, json
data = json.loads(aggregated_json_string)
# Merge daily rows back into a per-case format if needed,
# or call process_csv() after writing a temp CSV.
```

The script's `main()` function calls in order:
1. `process_csv(csv_path)` — builds the `data` dict
2. `build_excel(data, df_active, week_label)` — writes the 6-sheet Excel
3. `build_html(data)` — returns single-week HTML string (write it to file manually)
4. `load_history()` + `save_to_history(history, data)` — appends week to history
5. `build_html_v3(history)` — writes the multi-period HTML
6. Copies all 3 outputs to the permanent workspace

---

## Step 4 — Outputs

| File | Description |
|------|-------------|
| `Sure_Weekly_Dashboard_Mon_DD_to_Sun_DD_Mon_YYYY.xlsx` | 6-sheet Excel workbook |
| `Sure_Weekly_Dashboard.html` | Single-week visual dashboard |
| `Sure_Weekly_Dashboard_Multi.html` | Multi-period dashboard (year/month/week filters) |
| `dashboard_history.json` | Cumulative history — do not delete |

---

## Step 5 — History

`dashboard_history.json` accumulates every week's data. Key = Monday ISO date (e.g. `"2026-03-23"`).
The multi-period HTML reads all weeks from this file. Each Sunday run only adds/updates its own week.

Never overwrite the whole file — only `save_to_history()` should touch it.

---

## Targets reference (2026)

| Centre | Q1 Daily | Q1 Monthly | Q2 Daily | Q2 Monthly |
|--------|----------|------------|----------|------------|
| Ilasa Main Centre, Lagos | 200k | 5.0M | 228k | 5.7M |
| OAUTH Ilesa Centre | 220k | 5.5M | 280k | 7.0M |
| Palm Avenue, Lagos | 100k | 2.5M | 120k | 3.0M |
| Sure Ilesha, Osun (=Ijofi) | 170k | 4.25M | 210k | 5.25M |

---

## Common issues

**CDN blocked**: Use `mcp__Claude_in_Chrome__javascript_tool` to fetch — never `requests` from sandbox.

**UnicodeEncodeError surrogates**: Avoid emoji in JS string literals inside Python heredocs. Use plain ASCII labels.

**SyntaxError: identifier already declared**: Check for duplicate `const`/`let` names in the same function scope. Rename the duplicate.

**No CSV auto-detected**: Pass path explicitly or ensure filename contains "cases".
SKILLEOF

echo ""
echo "=== Skill installed successfully! ==="
echo "Location: $SKILL_DIR/SKILL.md"
echo ""
echo "Restart Cowork (or start a new session) to activate the skill."
echo ""
read -p "Press Enter to close..."
