"""
E-Commerce Data Analysis
========================
Script 01: Data Loading, Cleaning & KPI Analysis
Dataset : orders.csv (500 rows, Kaggle-style)
Run from: ecom_project/ folder
Command : python scripts/01_analysis.py
"""

import pandas as pd
import numpy as np

# ── 1. LOAD DATA ──────────────────────────────────────────────────────────────
df = pd.read_csv("data/orders.csv", parse_dates=["order_date", "ship_date"])

print("=" * 60)
print("  STEP 1 — RAW DATA SNAPSHOT")
print("=" * 60)
print(f"\n  Rows      : {df.shape[0]}")
print(f"  Columns   : {df.shape[1]}")
print(f"  Date range: {df['order_date'].min().date()} → {df['order_date'].max().date()}")
print(f"\n  Columns:\n  {list(df.columns)}")

# ── 2. DATA QUALITY CHECK ─────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("  STEP 2 — DATA QUALITY CHECK")
print("=" * 60)
print("\n  Missing values per column:")
print(df.isnull().sum().to_string())
print(f"\n  Duplicate rows : {df.duplicated().sum()}")
print(f"  Data types:\n{df.dtypes.to_string()}")

# ── 3. FEATURE ENGINEERING ────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("  STEP 3 — FEATURE ENGINEERING")
print("=" * 60)

df["month"]        = df["order_date"].dt.month_name()
df["month_num"]    = df["order_date"].dt.month
df["year"]         = df["order_date"].dt.year
df["quarter"]      = df["order_date"].dt.to_period("Q").astype(str)
df["ship_days"]    = (df["ship_date"] - df["order_date"]).dt.days
df["is_returned"]  = df["returned"].map({"Yes": 1, "No": 0})

print("\n  New columns added:")
print("  → month, month_num, year, quarter")
print("  → ship_days (delivery time in days)")
print("  → is_returned (1=Yes, 0=No for rate calculation)")

# ── 4. BUSINESS KPIs ──────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("  STEP 4 — BUSINESS KPIs")
print("=" * 60)

total_orders    = df.shape[0]
total_customers = df["customer_id"].nunique()
gross_revenue   = df["revenue"].sum()
total_discounts = df["discount_amt"].sum()
net_revenue     = df["net_revenue"].sum()
avg_order_val   = df["net_revenue"].mean()
return_rate     = df["is_returned"].mean() * 100
avg_rating      = df["rating"].mean()
avg_ship_days   = df["ship_days"].mean()

print(f"\n  Total Orders       : {total_orders:,}")
print(f"  Unique Customers   : {total_customers}")
print(f"  Gross Revenue      : ₹{gross_revenue:,.2f}")
print(f"  Total Discounts    : ₹{total_discounts:,.2f}")
print(f"  Net Revenue        : ₹{net_revenue:,.2f}")
print(f"  Avg Order Value    : ₹{avg_order_val:,.2f}")
print(f"  Return Rate        : {return_rate:.1f}%")
print(f"  Avg Customer Rating: {avg_rating:.2f} / 5")
print(f"  Avg Shipping Days  : {avg_ship_days:.1f} days")

# ── 5. REVENUE BY CATEGORY ────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("  STEP 5 — REVENUE BY CATEGORY")
print("=" * 60)

cat = (
    df.groupby("category")
    .agg(
        orders        = ("order_id",    "count"),
        net_revenue   = ("net_revenue", "sum"),
        avg_order     = ("net_revenue", "mean"),
        avg_rating    = ("rating",      "mean"),
        return_rate   = ("is_returned", "mean"),
    )
    .sort_values("net_revenue", ascending=False)
    .round(2)
)
cat["return_rate"] = (cat["return_rate"] * 100).round(1)
print("\n" + cat.to_string())

# ── 6. MONTHLY TREND ──────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("  STEP 6 — MONTHLY REVENUE TREND (2023)")
print("=" * 60)

monthly = (
    df[df["year"] == 2023]
    .groupby(["month_num", "month"])["net_revenue"]
    .sum()
    .reset_index()
    .sort_values("month_num")
    .set_index("month")["net_revenue"]
    .round(2)
)
print("\n" + monthly.to_string())

# ── 7. REGION ANALYSIS ────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("  STEP 7 — REVENUE BY REGION")
print("=" * 60)

region = (
    df.groupby("region")
    .agg(
        orders      = ("order_id",    "count"),
        net_revenue = ("net_revenue", "sum"),
        avg_order   = ("net_revenue", "mean"),
        avg_rating  = ("rating",      "mean"),
    )
    .sort_values("net_revenue", ascending=False)
    .round(2)
)
print("\n" + region.to_string())

# ── 8. TOP CUSTOMERS ──────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("  STEP 8 — TOP 10 CUSTOMERS BY LIFETIME VALUE")
print("=" * 60)

top_cust = (
    df.groupby(["customer_id", "customer_name"])
    .agg(
        orders          = ("order_id",    "count"),
        lifetime_value  = ("net_revenue", "sum"),
        avg_order       = ("net_revenue", "mean"),
        avg_rating      = ("rating",      "mean"),
    )
    .sort_values("lifetime_value", ascending=False)
    .head(10)
    .round(2)
)
print("\n" + top_cust.to_string())

# ── 9. SHIPPING ANALYSIS ──────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("  STEP 9 — SHIPPING MODE ANALYSIS")
print("=" * 60)

ship = (
    df.groupby("ship_mode")
    .agg(
        orders        = ("order_id",   "count"),
        avg_ship_days = ("ship_days",  "mean"),
        avg_rating    = ("rating",     "mean"),
        return_rate   = ("is_returned","mean"),
    )
    .round(2)
)
ship["return_rate"] = (ship["return_rate"] * 100).round(1)
print("\n" + ship.to_string())

# ── 10. PAYMENT METHOD ────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("  STEP 10 — PAYMENT METHOD USAGE")
print("=" * 60)

payment = (
    df.groupby("payment_method")
    .agg(orders=("order_id","count"), net_revenue=("net_revenue","sum"))
    .sort_values("orders", ascending=False)
    .round(2)
)
print("\n" + payment.to_string())

# ── SAVE CLEANED DATA ─────────────────────────────────────────────────────────
df.to_csv("data/orders_clean.csv", index=False)
print("\n" + "=" * 60)
print("  ✓ Cleaned data saved → data/orders_clean.csv")
print("=" * 60 + "\n")
