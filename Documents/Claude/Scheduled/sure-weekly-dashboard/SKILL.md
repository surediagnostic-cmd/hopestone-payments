---
name: sure-weekly-dashboard
description: Sure Finance weekly performance dashboard — export from LabSmartLIS, aggregate, build Excel + HTML
---

## Sure Finance Weekly Performance Dashboard (v3 — Multi-period)

**Objective:** Every Sunday at 11:59 PM WAT — export the week's data from LabSmartLIS, process it, append it to the history file, and regenerate both the single-week (v2) and the multi-period (v3) dashboards.

---

### Step 1 — Compute date range

Today is Sunday. Report period = **Mon–Sun of the week just ended**.
- `week_start` = today − 6 days (Monday), format `YYYY-MM-DD`
- `week_end`   = today (Sunday), format `YYYY-MM-DD`

---

### Step 2 — Export CSV from LabSmartLIS

Using Chrome MCP (`mcp__Claude_in_Chrome__*`):

1. Navigate to: `https://app.labsmartlis.com/72391405/data_exports/new`
2. Confirm Entity = **Case**; set date range `week_start` → `week_end`
3. Click **Run export** (use `javascript_tool` to submit if button click doesn't fire a network request)
4. Capture the CDN download URL via `read_network_requests` — format: `https://d33hv1py9n6skn.cloudfront.net/...`

> ⚠️ The CDN is **blocked from the sandbox**. Never try Python/curl to fetch it.

---

### Step 3 — Fetch CSV into browser memory

```javascript
const url = '<CDN_URL>';
const resp = await fetch(url);
window.__csvData = await resp.text();
`Fetched ${window.__csvData.length} chars, ${window.__csvData.split('\\n').length} rows`;
```

---

### Step 4 — Aggregate in JavaScript

Run this script in the same tab to compute all metrics and store in sessionStorage:

```javascript
const csv = window.__csvData;
const lines = csv.split('\n');
const headers = lines[0].split(',').map(h => h.replace(/^"|"$/g,'').trim());
function getIdx(name) { return headers.findIndex(h => h === name); }
const iDate=getIdx('Date'),iCentre=getIdx('Collection centre'),iCaseId=getIdx('Case Id'),
  iRegNo=getIdx('Reg. no.'),iInv=getIdx('Investigations'),iTotalFee=getIdx('Total Fee'),
  iFeePaid=getIdx('Fee Paid'),iFeeDue=getIdx('Fee Due'),iReferrer=getIdx('Referrer'),
  iAgent=getIdx('Booked by'),iStatus=getIdx('Status');
function countInv(v){if(!v||!v.trim())return 0;return Math.max(v.split(',').map(s=>s.trim()).filter(s=>s.length>0).length,1);}
function normAgent(n){return(!n||!n.trim())?'Walk-in / Unassigned':n.trim().replace(/\s+/g,' ');}
function normRef(n){return(!n||!n.trim())?'Walk-in / Self':n.trim().replace(/\s+/g,' ');}
const rows=[];
for(let i=1;i<lines.length;i++){
  const line=lines[i].trim();if(!line)continue;
  const fields=[];let inQ=false,cur='';
  for(let c=0;c<line.length;c++){if(line[c]==='"'){inQ=!inQ;}else if(line[c]===','&&!inQ){fields.push(cur);cur='';}else{cur+=line[c];}}
  fields.push(cur);
  if(fields[iStatus]&&fields[iStatus].trim().toLowerCase()==='cancelled')continue;
  rows.push(fields);
}
const dailyMap={};
for(const r of rows){
  const date=r[iDate]?r[iDate].trim():'',centre=r[iCentre]?r[iCentre].trim():'';
  const key=date+'|'+centre;
  if(!dailyMap[key])dailyMap[key]={date,centre,cases:new Set(),patients:new Set(),inv:0,rev:0,paid:0,due:0,refs:new Set()};
  const d=dailyMap[key];
  d.cases.add(r[iCaseId]);d.patients.add(r[iRegNo]);d.inv+=countInv(r[iInv]);
  d.rev+=parseFloat(r[iTotalFee])||0;d.paid+=parseFloat(r[iFeePaid])||0;d.due+=parseFloat(r[iFeeDue])||0;
  const ref=normRef(r[iReferrer]);if(ref!=='Walk-in / Self')d.refs.add(ref);
}
const daily=Object.values(dailyMap).map(d=>({date:d.date,centre:d.centre,cases:d.cases.size,patients:d.patients.size,inv:d.inv,rev:Math.round(d.rev),paid:Math.round(d.paid),due:Math.round(d.due),urefs:d.refs.size})).sort((a,b)=>a.date.localeCompare(b.date)||a.centre.localeCompare(b.centre));
const refMap={};
for(const r of rows){
  const centre=r[iCentre]?r[iCentre].trim():'',ref=normRef(r[iReferrer]);
  const key=centre+'|'+ref;if(!refMap[key])refMap[key]={centre,ref,cases:new Set(),rev:0};
  refMap[key].cases.add(r[iCaseId]);refMap[key].rev+=parseFloat(r[iTotalFee])||0;
}
const referrers=Object.values(refMap).map(d=>({centre:d.centre,ref:d.ref,cases:d.cases.size,rev:Math.round(d.rev)})).filter(d=>d.cases>0).sort((a,b)=>b.cases-a.cases);
const agentMap={};
for(const r of rows){
  const centre=r[iCentre]?r[iCentre].trim():'',agent=normAgent(r[iAgent]);
  const key=centre+'|'+agent;if(!agentMap[key])agentMap[key]={centre,agent,cases:new Set(),inv:0,rev:0};
  agentMap[key].cases.add(r[iCaseId]);agentMap[key].inv+=countInv(r[iInv]);agentMap[key].rev+=parseFloat(r[iTotalFee])||0;
}
const agents=Object.values(agentMap).map(d=>({centre:d.centre,agent:d.agent,cases:d.cases.size,inv:d.inv,rev:Math.round(d.rev)})).sort((a,b)=>b.cases-a.cases);
const allCases=new Set(rows.map(r=>r[iCaseId])),allPats=new Set(rows.map(r=>r[iRegNo]));
const allRefs=new Set(rows.map(r=>normRef(r[iReferrer])).filter(x=>x!=='Walk-in / Self'));
const totalRev=rows.reduce((s,r)=>s+(parseFloat(r[iTotalFee])||0),0);
const totalPaid=rows.reduce((s,r)=>s+(parseFloat(r[iFeePaid])||0),0);
const totalDue=rows.reduce((s,r)=>s+(parseFloat(r[iFeeDue])||0),0);
sessionStorage.setItem('daily',JSON.stringify(daily));
sessionStorage.setItem('referrers',JSON.stringify(referrers));
sessionStorage.setItem('agents',JSON.stringify(agents));
sessionStorage.setItem('summary',JSON.stringify({totalRev:Math.round(totalRev),totalPaid:Math.round(totalPaid),totalDue:Math.round(totalDue),totalCases:allCases.size,totalPats:allPats.size,totalRefs:allRefs.size,totalInv:daily.reduce((s,r)=>s+r.inv,0),cancelled:lines.length-1-rows.length-1}));
`daily:${daily.length} refs:${referrers.length} agents:${agents.length} cases:${allCases.size}`;
```

---

### Step 5 — Retrieve data in chunks

Retrieve each JSON array from sessionStorage in slices of ~10 rows to avoid truncation:

```javascript
JSON.stringify(JSON.parse(sessionStorage.getItem('daily')).slice(0, 10))
// then .slice(10, 20) etc until all rows retrieved
JSON.stringify(JSON.parse(sessionStorage.getItem('summary')))
JSON.stringify(JSON.parse(sessionStorage.getItem('referrers')).slice(0, 15))
JSON.stringify(JSON.parse(sessionStorage.getItem('agents')))
```

Assemble all rows into Python-side lists.

---

### Step 6 — Build `data` dict and call pipeline

Key name mappings (JS field → Python `data` dict):
- `daily[*].date` → `Date`
- `daily[*].centre` → `Collection centre`
- `daily[*].inv` → `investigations`
- `daily[*].rev` → `revenue`
- `daily[*].urefs` → `unique_referrers`
- `referrers[*].ref` → `Referrer`
- `referrers[*].centre` → `Collection centre`
- `agents[*].agent` → `Agent`

Find the pipeline script at:
- `/sessions/[session-id]/mnt/Sure Finance & Business Scorecards/update_dashboard.py`

Import it and call:

```python
import importlib.util, json, datetime, shutil, os, pandas as pd

spec = importlib.util.spec_from_file_location('upd', '<path_to_update_dashboard.py>')
upd = importlib.util.module_from_spec(spec)
spec.loader.exec_module(upd)

# ... build data dict from aggregated values (weekly totals from daily sums) ...
# ... build minimal df_active DataFrame ...

xlsx_path = upd.build_excel(data, df_active, week_label)
html_content = upd.build_html(data)

# Single-week HTML
html_path = '/sessions/[session-id]/Sure_Weekly_Dashboard.html'
with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html_content)

# History + multi-period HTML (v3)
history = upd.load_history()
history = upd.save_to_history(history, data)
multi_path = upd.build_html_v3(history)
```

---

### Step 7 — Save to workspace

Copy all outputs to: `/sessions/[session-id]/mnt/Sure Finance & Business Scorecards/`

Files to save:
- `Sure_Weekly_Dashboard_<week_label>.xlsx` — Excel (6 sheets)
- `Sure_Weekly_Dashboard.html` — current week HTML (v2, overwrite)
- `Sure_Weekly_Dashboard_Multi.html` — multi-period dashboard (v3, overwrite)
- `dashboard_history.json` — updated automatically by `save_to_history()`

`update_dashboard.py` also automatically calls `update_artifact_html(data)` which injects
the new week into both live sidebar dashboards:
- `/Users/mac/Documents/Claude/Artifacts/sure-weekly-dashboard/index.html`
- `/Users/mac/Documents/01_Sure_Diagnostics/Dashboards/Sure Labsmart Dashboard/weekly/index.html`

No manual copy step needed — the sidebar previews update as part of the pipeline.

---

### Notes

- **Timezone**: WAT (UTC+1, Africa/Lagos). All dates in WAT.
- **Sure Ilesha, Osun** in LabSmartLIS = **Ijofi branch** in the targets table. Key in TARGETS dict is `Sure Ilesha, Osun`.
- **Targets**: Q1=Jan–Mar, Q2=Apr–Jun, Q3=Jul–Sep, Q4=Oct–Dec. Embedded in update_dashboard.py TARGETS dict.
- **History file** persists at: `mnt/Sure Finance & Business Scorecards/dashboard_history.json`. `save_to_history()` appends to it. `load_history()` reads it.
- **Multi-period dashboard** (`_Multi.html`) shows year/month/week filters + Summary (aggregated) / Weekly drill-down toggle.

### Success criteria

- Excel saved with 6 sheets
- `Sure_Weekly_Dashboard.html` updated (current week)
- `Sure_Weekly_Dashboard_Multi.html` updated with new week added to history
- `dashboard_history.json` updated with new week's key
- History count increases by 1 each week
