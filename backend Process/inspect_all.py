import pandas as pd
import json

results = {}

# ─── Domestic Lounge ───────────────────────────────────────────────
xl = pd.ExcelFile(r"c:\Lounges report\Domestic Lounge.xlsx")
results["Domestic_sheets"] = xl.sheet_names
for sh in xl.sheet_names:
    df = xl.parse(sh, nrows=2)
    results[f"Domestic_{sh}_cols"] = list(df.columns)

# ─── Railway Lounge ────────────────────────────────────────────────
xl2 = pd.ExcelFile(r"c:\Lounges report\Railway Lounge.xlsx")
results["Railway_sheets"] = xl2.sheet_names
for sh in xl2.sheet_names:
    df = xl2.parse(sh, nrows=2)
    results[f"Railway_{sh}_cols"] = list(df.columns)

# ─── Global Lounge ─────────────────────────────────────────────────
xl3 = pd.ExcelFile(r"c:\Lounges report\Global Lounge.xlsx")
results["Global_sheets"] = xl3.sheet_names
for sh in xl3.sheet_names:
    df = xl3.parse(sh, nrows=3)
    results[f"Global_{sh}_cols"] = list(df.columns)

# ─── Data.csv ──────────────────────────────────────────────────────
df_data = pd.read_csv(r"c:\Lounges report\Data.csv", nrows=3)
results["Data_cols"] = list(df_data.columns)
results["Data_service_names"] = list(pd.read_csv(r"c:\Lounges report\Data.csv", usecols=["service_name"])["service_name"].unique())
results["Data_partner_names_sample"] = list(pd.read_csv(r"c:\Lounges report\Data.csv", usecols=["partner_name"])["partner_name"].unique()[:30])

print(json.dumps(results, indent=2, default=str))
