"""Verify (Daily) hub filter, lots-based aggregation, VWAP."""
import base64
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import app_pt

path = Path(__file__).parent / "Futs Portfolio.xls.xlsm"
raw = path.read_bytes()
contents = "data:application/vnd.ms-excel;base64," + base64.b64encode(raw).decode()

df, err = app_pt.parse_uploaded_file(contents, path.name)
print("parse error:", err)
print("rows kept:", len(df))
print()
print("Unique hubs after filter — none should contain (Daily):")
for h in sorted(df["Hub"].unique()):
    print(" ", repr(h))
print()
print("Signed Lots column present:", "Signed Lots" in df.columns)
print()
print("=== SP15 DA Peak Aug26 sanity check ===")
mask = (
    (df["Hub"] == "SP15 DA")
    & (df["Product"] == "Peak Futures")
    & (df["Delivery Month"] == "2026-08-01")
)
focus = df.loc[mask, ["Trade Date", "B/S", "Lots", "Total Quantity", "Price"]]
print(focus.to_string(index=False))
print(f"\nNet Lots (Aug26 SP15 DA Peak): {df.loc[mask, 'Signed Lots'].sum():+.0f}")
print(f"Trade VWAP: ${app_pt._weighted_avg(df.loc[mask, 'Price'], df.loc[mask, 'Total Quantity']):.2f}/MWh")
print()
print("=== Hub × Month aggregation (head) ===")
hub_month = app_pt.aggregate_monthly(df, ["Hub"])
print(hub_month.head(10).to_string(index=False))
print()
print("=== Position summary with Lots + VWAP (Aug 2026 rows) ===")
summary = app_pt.build_position_summary(df)
aug = summary[summary["Delivery Month"] == "2026-08-01"]
cols = ["Hub", "Product", "Net Lots", "Avg Buy ($/MWh)", "Avg Sell ($/MWh)", "Trade VWAP ($/MWh)"]
print(aug[cols].to_string(index=False))
