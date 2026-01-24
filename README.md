# Samsung
Process Optimization & Automation Toolkit

This repository contains internal automation tools built to streamline recurring SCM / Planning / Order Management workflows, improving efficiency, accuracy, and team coverage.

## Contents

### Python (Jupyter Notebooks)

**1) Mon ATS — Weekly Available To Stock (ATS) Report**
- Schedule: Every Monday morning (HTV & LFD)
- Input: GSCM export, SAP export, and weekly email files (manual download)
- Output: Standardized ATS Excel report (SKU × account)
- What it does: data cleaning, formatting, aggregation, ATS calculation, report write-back
- Runtime: ~10–15 min
- Time saved: ~2–3 hours/week (prep time ~5–10 min)

**2) Buffer File — Forecast Support (SCM Day / MKT)**
- Schedule: Every Thursday before SCM Day
- Users: SCM + Sales/Marketing (external stakeholders)
- Output: 7 AM Excel report supporting 6-month planning (SKU × account)
- Includes: AP1, RTF, calculated AP2, WoW AP1, product info, latest forecast, forecast guide, WOS, ending OH, adjusted sell-through/WOS, price
- Time saved: typically reduces 2–3 hours of manual work

**3) Problem Order Scanner — Order Exception Identification**
- Use: ad-hoc / daily
- Purpose: quickly flags orders that require attention (e.g., split requirement, pending ship, credit hold scenarios)
- Input: SAP download
- Output: filtered/flagged order list for follow-up with responsible teams

---

### VBA (Excel + SAP)

**4) Display SCH.Line Drop Tools (SAP schedule line update)**
Files:
- `Display SCH.Line-Drop click.bas`
- `Display SCH.Line-SAP function.bas`

How it works:
- Populate a template with `SO#`, `Line`, `Sch.line`
- Click **Drop**
- VBA updates SAP schedule lines and returns a status message per order

Impact:
- Helps accelerate a high-frequency “drop order” workflow where schedule line updates are required
- Saves ~20–30% time for daily order drop (requires maintenance after SAP UI updates)

## Tech Stack (key libraries)
- Python: `pandas`, `numpy`, `openpyxl`, `win32com.client`
- Date/time: `datetime`, `dateutil.relativedelta`
- Utilities: `os`, `time`, `copy`, `warnings`

(Other imports may appear in notebooks based on specific report needs.)
