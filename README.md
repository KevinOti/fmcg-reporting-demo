# FMCG Automated Reporting Pipeline (Demo)

**What this is:** a working demonstration of the reporting automation I build for
businesses — using entirely fictional data for a made-up beverage company,
"Savanna Beverages". No real company data appears anywhere in this project.

**The problem it solves:** every month, sales data arrives from multiple
distributors in different formats — different column names, mixed date formats,
typo'd product names, duplicate rows, junk header rows. Someone spends hours
cleaning and consolidating it in Excel before management can see a single number.

**What this pipeline does in one run (seconds, not hours):**

1. **Collect** — reads three distributor files in three different formats
   (tidy CSV, human-maintained Excel, and a system export with its own schema)
2. **Clean** — removes duplicates, skips junk headers, standardises mixed date
   formats, maps typo'd product names to a canonical catalogue, flags missing data
3. **Report** — builds a formatted Excel workbook:
   - Sales vs Target by region (with achievement %)
   - Product performance ranking
   - Distributor summary
   - **Data Quality Log** — every issue caught and fixed, as an audit trail
4. **Deliver** — emails the report through Outlook on a schedule
   (Windows Task Scheduler), so it arrives before anyone is at their desk

## Run it yourself

```
pip install pandas openpyxl
python generate_data.py   # creates the messy raw files in /raw
python pipeline.py        # cleans, consolidates, writes /output report
```

To enable the email step on Windows: `pip install pywin32`, set
`SEND_EMAIL = True` and a recipient in `pipeline.py`, and schedule
`pipeline.py` with Task Scheduler.

## Stack

Python (pandas, openpyxl) · SQL Server in production versions ·
Windows Task Scheduler · Outlook (win32com) · Power BI for dashboards

---
*Kevin Otieno — [kevinoti.github.io/portfolio](https://kevinoti.github.io/portfolio/) · kevinotieno.ot@gmail.com*
