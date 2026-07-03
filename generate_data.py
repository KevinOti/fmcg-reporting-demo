"""
generate_data.py — creates realistic, MESSY dummy FMCG sales data.

Simulates what actually lands in an analyst's inbox: three distributors,
three different file formats, three different kinds of mess.
Run: python generate_data.py
"""
import random
import numpy as np
import pandas as pd
from pathlib import Path

random.seed(42)
np.random.seed(42)

RAW = Path(__file__).parent / "raw"
RAW.mkdir(exist_ok=True)

# --- Fictional FMCG product catalogue (beverage company) ---
SKUS = [
    ("SAV-TEA-050", "Savanna Chai 50g"),
    ("SAV-TEA-100", "Savanna Chai 100g"),
    ("SAV-TEA-250", "Savanna Chai 250g"),
    ("SAV-COF-100", "Savanna Coffee 100g"),
    ("SAV-COF-250", "Savanna Coffee 250g"),
    ("SAV-COC-500", "Savanna Cocoa 500g"),
    ("SAV-JUI-1LT", "Savanna Juice 1L"),
    ("SAV-JUI-500", "Savanna Juice 500ml"),
]

REGIONS = ["Nairobi", "Central", "Coast", "Western", "Rift Valley"]

DAYS = pd.date_range("2026-06-01", "2026-06-30", freq="D")

def base_rows(distributor, regions, date_fmt, sku_style, n_outlets):
    """Generate clean-ish rows, then each distributor's mess is layered on top."""
    rows = []
    for day in DAYS:
        for region in regions:
            for code, name in SKUS:
                if random.random() < 0.35:   # not every SKU sells everywhere daily
                    continue
                units = max(1, int(np.random.gamma(2.2, 14)))
                price = {"050": 85, "100": 160, "250": 370, "500": 250, "1LT": 180}[code[-3:]]
                rows.append({
                    "Date": day.strftime(date_fmt),
                    "Distributor": distributor,
                    "Region": region,
                    "SKU": code if sku_style == "code" else name,
                    "Units": units,
                    "UnitPrice": price,
                    "Outlets": random.randint(1, n_outlets),
                })
    return pd.DataFrame(rows)

# ---------------------------------------------------------------------------
# Distributor A — "Pwani Distributors" (CSV, mostly tidy but with duplicates
# and some negative/zero unit entries from their system's returns handling)
# ---------------------------------------------------------------------------
df_a = base_rows("Pwani Distributors", ["Coast", "Nairobi"], "%Y-%m-%d", "code", 8)
dupes = df_a.sample(frac=0.04, random_state=1)                # duplicate rows
df_a = pd.concat([df_a, dupes]).sample(frac=1, random_state=2).reset_index(drop=True)
df_a.loc[df_a.sample(frac=0.01, random_state=3).index, "Units"] = 0
df_a.to_csv(RAW / "pwani_daily_sales_june.csv", index=False)

# ---------------------------------------------------------------------------
# Distributor B — "Highlands Agencies" (Excel, human-maintained: junk header
# rows, inconsistent date formats, SKU names typed by hand with typos/case,
# missing values, a stray TOTAL row)
# ---------------------------------------------------------------------------
df_b = base_rows("Highlands Agencies", ["Central", "Rift Valley"], "%d/%m/%Y", "name", 12)
name_mess = {
    "Savanna Chai 50g":  ["savanna chai 50g", "Savana Chai 50g", "SAVANNA CHAI 50G", "Savanna Chai 50 g"],
    "Savanna Chai 100g": ["Savanna chai 100g", "Savanna Chai 100gm"],
    "Savanna Coffee 100g": ["Sav Coffee 100g", "savanna coffee 100g"],
    "Savanna Juice 1L":  ["Savanna Juice 1 Litre", "Savanna Juice 1Lt"],
}
def messify(name):
    if name in name_mess and random.random() < 0.5:
        return random.choice(name_mess[name])
    return name
df_b["SKU"] = df_b["SKU"].apply(messify)
df_b.loc[df_b.sample(frac=0.03, random_state=4).index, "Units"] = np.nan
# random date-format switching mid-file (classic human data entry)
def flip_date(d):
    if random.random() < 0.25:
        day, month, year = d.split("/")
        return f"{int(day)}-{int(month)}-{year[2:]}"     # 3-6-26
    return d
df_b["Date"] = df_b["Date"].apply(flip_date)

from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.title = "June"
ws.append(["HIGHLANDS AGENCIES LTD"])           # junk row 1
ws.append(["Sales Return - June 2026"])          # junk row 2
ws.append([])                                    # junk row 3
ws.append(list(df_b.columns))
for r in df_b.itertuples(index=False):
    ws.append(list(r))
ws.append(["TOTAL", "", "", "", float(df_b["Units"].sum(skipna=True)), "", ""])  # stray total row
wb.save(RAW / "highlands_june_2026.xlsx")

# ---------------------------------------------------------------------------
# Distributor C — "Lakeview Traders" (CSV with DIFFERENT column names and
# column order — the "every system exports differently" problem)
# ---------------------------------------------------------------------------
df_c = base_rows("Lakeview Traders", ["Western", "Nairobi"], "%m/%d/%Y", "code", 6)
df_c = df_c.rename(columns={
    "Date": "txn_date", "SKU": "product_code", "Units": "qty_sold",
    "UnitPrice": "price_per_unit", "Region": "sales_region",
    "Distributor": "dist_name", "Outlets": "outlet_count",
})
df_c = df_c[["dist_name", "txn_date", "product_code", "qty_sold",
             "price_per_unit", "sales_region", "outlet_count"]]
df_c.to_csv(RAW / "lakeview_export_jun26.csv", index=False)

# ---------------------------------------------------------------------------
# Monthly targets file (what management set per region)
# ---------------------------------------------------------------------------
targets = pd.DataFrame({
    "Region": REGIONS,                       # Nairobi, Central, Coast, Western, Rift Valley
    "MonthlyTargetKsh": [2_300_000, 850_000, 1_150_000, 1_000_000, 1_500_000],
})
targets.to_csv(RAW / "june_targets.csv", index=False)

print("Raw files generated:")
for f in sorted(RAW.iterdir()):
    print("  ", f.name)
