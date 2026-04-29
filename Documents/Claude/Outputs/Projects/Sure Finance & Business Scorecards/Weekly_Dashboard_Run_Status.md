# Sure Finance Weekly Dashboard — Automated Run Status

**Run Date:** Monday, 6 April 2026 (scheduled for Sunday 5 April 2026 at 23:59 WAT)
**Report Period:** Monday 30 March – Sunday 5 April 2026
**Status:** ⚠️ BLOCKED — LabSmartLIS Session Expired

---

## What Happened

The automated weekly dashboard task ran as scheduled but could not complete because the LabSmartLIS session has expired. When navigating to:

```
https://app.labsmartlis.com/72391405/data_exports/new
```

The system redirected to the login page (`/labsmart_sessions/new`), requiring email and password authentication.

For security reasons, the automated task cannot enter account passwords. The user must log in manually.

---

## How to Run the Dashboard Manually

1. **Log in** to LabSmartLIS at https://app.labsmartlis.com
2. **Open a new chat** in Cowork
3. **Ask Claude** to run the weekly dashboard for the week of **30 March – 5 April 2026**

Claude will:
- Navigate to the export page (already logged in via your browser)
- Export the CSV for that date range
- Process and generate the Excel + HTML dashboards
- Save them to your workspace folder

---

## Previous Dashboard (for reference)

The last successfully generated weekly dashboard covers:
- **Week:** Mon 23 Mar – Sun 29 Mar 2026
- **Files:** `Sure_Weekly_Dashboard_Mon_23_to_Sun_29_Mar_2026.xlsx` and `.html`
- **Multi-period:** `Sure_Weekly_Dashboard_Multi.html`

---

*Generated automatically by the Sure Finance Weekly Dashboard scheduled task.*
