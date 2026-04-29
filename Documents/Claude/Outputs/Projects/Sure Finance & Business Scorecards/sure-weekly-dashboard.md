---
name: sure-weekly-dashboard
description: >
  Run the Sure Finance & Business weekly performance dashboard pipeline. Use this skill whenever
  the user wants to generate the weekly dashboard, process LabSmartLIS data, update performance
  reports, run the weekly pipeline, or refresh the Sure Finance dashboard. Also triggers when the
  user mentions exporting cases from LabSmartLIS, updating dashboard history, or generating the
  Excel/HTML performance reports.
---

# Sure Finance Weekly Dashboard Pipeline

Generates the weekly performance dashboard from LabSmartLIS data: Excel workbook (6 sheets),
single-week HTML dashboard, and multi-period HTML dashboard with year/month/week filters.

## Key constraints

- **Pipeline script**: `update_dashboard.py` in the session working directory
- **Permanent outputs**: the "Sure Finance & Business Scorecards" workspace folder
- **History file**: `dashboard_history.json` in that folder — persists across sessions, do not delete
- **CDN blocked**: LabSmartLIS downloads from `d33hv1py9n6skn.cloudfront.net`, which the Python
  sandbox CANNOT reach. Always use `mcp__Claude_in_Chrome__javascript_tool` to fetch via browser.
- **Centre mapping**: "Sure Ilesha, Osun" = Ijofi branch targets in the script
- **Timezone**: WAT (UTC+1). Scheduled at cron `59 23 * * 0` = 11:59 PM WAT every Sunday.

## Step 1 — Export CSV from LabSmartLIS

Open Chrome at https://ng.labsmart.net → Business → Data Export
- Entity: Case, Date range: Mon 00:00 to Sun 23:59 for the target week
- Click Export. Copy the CDN download URL — do NOT fetch it from Python.

## Step 2 — Fetch and aggregate in-browser
```javascript
// Run via mcp__Claude_in_Chrome__javascript_tool in the LabSmartLIS tab
// First get the CDN URL from the page:
Array.from(document.querySelectorAll("a[href*=\"cloudfront\"]")).map(a=>a.href)
```

Then run the aggregation script (fetches CSV via browser, aggregates daily x centre, stores in sessionStorage):
```javascript
(async()=>{
  const url="PASTE_CDN_URL";
  const text=await(await fetch(url)).text();
  const lines=text.split("\n");
  const H=lines[0].split(",").map(h=>h.trim().replace(/^"|"$/g,""));
  function col(r,n){const i=H.indexOf(n);return i>=0?(r[i]||"").replace(/^"|"$/g,"").trim():"";}
  const rows=[];
  for(let i=1;i<lines.length;i++){const r=lines[i].match(/(".*?"|[^,]+)(?=,|$)/g);if(r&&r.length>=5)rows.push(r);}
  const D={},R={},A={};
  for(const r of rows){
    const c=col(r,"Canceled").toLowerCase();if(c==="true"||c==="1")continue;
    const dt=col(r,"Date").split(" ")[0],cn=col(r,"Collection centre"),k=dt+"|"+cn;
    const rev=parseFloat(col(r,"Total Fee"))||0,paid=parseFloat(col(r,"Fee Paid"))||0,due=parseFloat(col(r,"Fee Due"))||0;
    const inv=col(r,"Investigations"),ic=inv?inv.split(",").filter(x=>x.trim()).length||1:0;
    const pt=col(r,"Reg. no."),ref=col(r,"Referrer"),ag=col(r,"Agent")||"Walk-in / Unassigned";
    if(!D[k])D[k]={date:dt,centre:cn,inv:0,cases:0,patients:new Set(),rev:0,paid:0,due:0,refs:new Set()};
    D[k].inv+=ic;D[k].cases++;D[k].patients.add(pt);D[k].rev+=rev;D[k].paid+=paid;D[k].due+=due;D[k].refs.add(ref);
    const rk=cn+"|"+ref;if(!R[rk])R[rk]={centre:cn,referrer:ref,cases:0,inv:0,rev:0};R[rk].cases++;R[rk].inv+=ic;R[rk].rev+=rev;
    if(!A[ag])A[ag]={agent:ag,cases:0,inv:0,patients:new Set(),rev:0,paid:0};A[ag].cases++;A[ag].inv+=ic;A[ag].patients.add(pt);A[ag].rev+=rev;A[ag].paid+=paid;
  }
  const da=Object.values(D).map(d=>({"Date":d.date,"Collection centre":d.centre,investigations:d.inv,cases:d.cases,patients:d.patients.size,revenue:Math.round(d.rev),paid:Math.round(d.paid),due:Math.round(d.due),unique_referrers:d.refs.size}));
  const ra=Object.values(R).map(r=>({"Collection centre":r.centre,Referrer:r.referrer,cases:r.cases,investigations:r.inv,revenue:Math.round(r.rev)})).sort((a,b)=>b.cases-a.cases);
  const aa=Object.values(A).map(a=>({Agent:a.agent,cases:a.cases,investigations:a.inv,patients:a.patients.size,revenue:Math.round(a.rev),paid:Math.round(a.paid)})).sort((a,b)=>b.investigations-a.investigations);
  const s=JSON.stringify({daily:da,referrers:ra,agents:aa});
  const C=200000,n=Math.ceil(s.length/C);
  sessionStorage.setItem("dash_nchunks",n);
  for(let i=0;i<n;i++)sessionStorage.setItem("dash_chunk_"+i,s.slice(i*C,(i+1)*C));
  return "Stored "+n+" chunks, "+da.length+" daily rows, "+ra.length+" referrers, "+aa.length+" agents";
})();
```

Retrieve chunks:
```javascript
// Chunk count
parseInt(sessionStorage.getItem("dash_nchunks")||"0")
// Each chunk (repeat for _1, _2...)
sessionStorage.getItem("dash_chunk_0")
```

## Step 3 — Run the pipeline
```bash
python3 update_dashboard.py /path/to/cases.csv
```

This calls: process_csv → build_excel → build_html → load_history → save_to_history → build_html_v3
Then copies Excel, single-week HTML, and multi-period HTML to the permanent workspace.

## Step 4 — Outputs

| File | Description |
|------|-------------|
| `Sure_Weekly_Dashboard_Mon_DD_to_Sun_DD_Mon_YYYY.xlsx` | 6-sheet Excel |
| `Sure_Weekly_Dashboard.html` | Single-week dashboard |
| `Sure_Weekly_Dashboard_Multi.html` | Multi-period with year/month/week filters |
| `dashboard_history.json` | Cumulative history — never delete |

## Targets 2026

| Centre | Q1 Daily | Q1 Monthly |
|--------|----------|------------|
| Ilasa Main Centre, Lagos | 200k | 5.0M |
| OAUTH Ilesa Centre | 220k | 5.5M |
| Palm Avenue, Lagos | 100k | 2.5M |
| Sure Ilesha, Osun (=Ijofi) | 170k | 4.25M |

## Common issues

- **CDN blocked**: Use Chrome javascript_tool, never requests/urllib from sandbox
- **UnicodeEncodeError surrogates**: No emoji in JS string literals inside Python heredocs
- **SyntaxError identifier already declared**: Rename duplicate const/let in same function scope
- **No CSV found**: Pass path explicitly or ensure filename contains "cases"