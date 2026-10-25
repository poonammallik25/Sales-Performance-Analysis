"""
Sales Performance - End-to-End Analysis (Capstone, Week 4)
Step 1: Clean & explore  |  Step 2: Analyse (Pandas)  |  Step 3: Charts for BI/report
Run from the project root:  python 02_Code/01_clean_and_analyze.py
"""
import json, os
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mt
import seaborn as sns

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(ROOT, "01_Data", "Sample_data_raw.xlsx")
CLEAN = os.path.join(ROOT, "01_Data", "sales_clean.csv")
CHARTS = os.path.join(ROOT, "03_Dashboard", "charts")
os.makedirs(CHARTS, exist_ok=True)

# ------------------------------------------------------------------ 1. LOAD
df = pd.read_excel(RAW)
print("Raw shape:", df.shape)

# ------------------------------------------------------------------ 2. CLEAN
df.columns = df.columns.str.strip().str.replace(" ", "_")           # ' Sales' had a leading space
df["Discount_Band"] = df["Discount_Band"].fillna("No Discount")             # blank = no discount (Discounts == 0)
assert (df.loc[df.Discount_Band == "No Discount", "Discounts"] == 0).all()
df["Date"] = pd.to_datetime(df["Date"])
df["Month_Name"] = df["Month_Name"].str.strip()

# validation of business rules
assert (df.Gross_Sales - df.Discounts - df.Sales).abs().max() < 0.01   # Gross - Discounts = Sales
assert (df.Sales - df.COGS - df.Profit).abs().max() < 0.01             # Sales - COGS = Profit
print("Duplicates:", df.duplicated().sum(), "| Nulls left:", int(df.isna().sum().sum()))

# feature engineering
df["Profit_Margin_Pct"] = (df.Profit / df.Sales * 100).round(2)
df["Discount_Pct"] = (df.Discounts / df.Gross_Sales * 100).round(2)
df["Year_Month"] = df.Date.dt.strftime("%Y-%m")
df["Quarter"] = "Q" + df.Date.dt.quarter.astype(str)
df["Loss_Making"] = (df.Profit < 0).astype(int)

# data-quality notes (documented in the report)
df["COGS_per_Unit"] = (df.COGS / df.Units_Sold).round(2)
dq = {
    "rows": len(df),
    "half_unit_rows": int((df.Units_Sold % 1 != 0).sum()),
    "blank_discount_band_filled_none": 53,
    "mfg_price_matches_cogs_per_unit_rows": int((df.Manufacturing_Price == df.COGS_per_Unit).sum()),
    "date_range": [str(df.Date.min().date()), str(df.Date.max().date())],
}
df.to_csv(CLEAN, index=False)
print("Saved", CLEAN)

# ------------------------------------------------------------------ 3. ANALYSE
def summarize(by):
    g = df.groupby(by).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"),
                           Units=("Units_Sold", "sum"), Discounts=("Discounts", "sum"),
                           Gross=("Gross_Sales", "sum"), Orders=("Sales", "size"))
    g["Margin_%"] = (g.Profit / g.Sales * 100).round(1)
    g["Avg_Discount_%"] = (g.Discounts / g.Gross * 100).round(1)
    return g.sort_values("Sales", ascending=False)

seg, ctry, prod, band = (summarize(c) for c in ["Segment", "Country", "Product", "Discount_Band"])
month = summarize("Year_Month").sort_index()
like = df[df.Date.dt.month >= 9].groupby("Year").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"))
seg_band_margin = (df.groupby(["Segment", "Discount_Band"])[["Profit", "Sales"]].sum()
                   .assign(m=lambda x: x.Profit / x.Sales * 100)["m"].unstack().round(1))
ent = df[df.Segment == "Enterprise"]

kpi = {
    "total_sales": float(df.Sales.sum()), "total_profit": float(df.Profit.sum()),
    "margin_pct": float(df.Profit.sum() / df.Sales.sum() * 100),
    "units": float(df.Units_Sold.sum()), "discounts": float(df.Discounts.sum()),
    "gross": float(df.Gross_Sales.sum()), "orders": int(len(df)),
    "loss_rows": int(df.Loss_Making.sum()), "loss_value": float(df.loc[df.Profit < 0, "Profit"].sum()),
    "yoy_sales_pct": float((like.loc[2014, "Sales"] / like.loc[2013, "Sales"] - 1) * 100),
    "yoy_profit_pct": float((like.loc[2014, "Profit"] / like.loc[2013, "Profit"] - 1) * 100),
    "like_2013_sales": float(like.loc[2013, "Sales"]), "like_2014_sales": float(like.loc[2014, "Sales"]),
    "ent_discounts": float(ent.Discounts.sum()), "ent_profit": float(ent.Profit.sum()),
    "ent_sales": float(ent.Sales.sum()),
    "ent_gross_margin_before_disc_pct": float((ent.Gross_Sales.sum() - ent.COGS.sum()) / ent.Gross_Sales.sum() * 100),
    "dq": dq,
}
for name, t in {"segment": seg, "country": ctry, "product": prod, "discount_band": band, "month": month}.items():
    t.to_csv(os.path.join(ROOT, "03_Dashboard", f"summary_by_{name}.csv"))
json.dump(kpi, open(os.path.join(ROOT, "03_Dashboard", "kpis.json"), "w"), indent=2)
print(json.dumps({k: (round(v, 2) if isinstance(v, float) else v) for k, v in kpi.items()}, indent=1))
print(seg, "\n", band, "\n", seg_band_margin)

# ------------------------------------------------------------------ 4. CHARTS
NAVY, TEAL, ORANGE, RED, GREY = "#1F3A5F", "#2A9D8F", "#E9A23B", "#D1495B", "#9AA5B1"
plt.rcParams.update({"font.family": "DejaVu Sans", "axes.spines.top": False, "axes.spines.right": False,
                     "axes.titleweight": "bold", "axes.titlesize": 13, "figure.dpi": 150})
M = mt.FuncFormatter(lambda x, _: f"${x/1e6:.0f}M" if abs(x) >= 1e6 else f"${x/1e3:.0f}K")

def save(fig, name):
    fig.tight_layout(); fig.savefig(os.path.join(CHARTS, name), bbox_inches="tight"); plt.close(fig)

# 1 Sales & profit by segment
fig, ax = plt.subplots(figsize=(8, 4.6)); s = seg.reset_index(); x = np.arange(len(s))
ax.bar(x - .2, s.Sales, .4, color=NAVY, label="Sales"); ax.bar(x + .2, s.Profit, .4, color=[TEAL if v >= 0 else RED for v in s.Profit], label="Profit")
ax.set_xticks(x); ax.set_xticklabels(s.Segment, rotation=15); ax.yaxis.set_major_formatter(M); ax.axhline(0, color="k", lw=.6)
for i, m in enumerate(s["Margin_%"]): ax.text(i + .2, max(s.Profit.iloc[i], 0) + 1e6, f"{m}%", ha="center", fontsize=9, fontweight="bold")
ax.set_title("Sales vs Profit by Segment (labels = profit margin)"); ax.legend(frameon=False); save(fig, "01_segment.png")

# 2 Monthly trend
fig, ax = plt.subplots(figsize=(9, 4.4)); ax.plot(month.index, month.Sales, marker="o", color=NAVY, label="Sales")
ax.plot(month.index, month.Profit, marker="o", color=TEAL, label="Profit"); ax.yaxis.set_major_formatter(M)
ax.set_xticks(range(len(month))); ax.set_xticklabels(month.index, rotation=60, fontsize=8)
for k in ["2014-10", "2014-12"]:
    ax.annotate("Peak", (list(month.index).index(k), month.loc[k, "Sales"]), textcoords="offset points", xytext=(0, 8), ha="center", fontsize=8, color=ORANGE, fontweight="bold")
ax.set_title("Monthly Sales & Profit (Sep 2013 - Dec 2014)"); ax.legend(frameon=False); save(fig, "02_monthly_trend.png")

# 3 Discount band margin
fig, ax = plt.subplots(figsize=(7, 4.4)); order = ["No Discount", "Low", "Medium", "High"]; b = band.loc[order]
bars = ax.bar(order, b["Margin_%"], color=[TEAL, "#5BB5A8", ORANGE, RED])
for r, v in zip(bars, b["Margin_%"]): ax.text(r.get_x() + r.get_width() / 2, v + .3, f"{v}%", ha="center", fontweight="bold")
ax.set_ylabel("Profit margin %"); ax.set_title("Deeper Discounts = Thinner Margin"); save(fig, "03_discount_margin.png")

# 4 Country
fig, ax = plt.subplots(figsize=(7.5, 4.4)); c = ctry.sort_values("Profit", ascending=False).reset_index(); c["Country"] = c.Country.replace({"United States of America": "USA"})
ax.barh(c.Country[::-1], c.Profit[::-1], color=NAVY); ax.xaxis.set_major_formatter(M)
for i, (p, m) in enumerate(zip(c.Profit[::-1], c["Margin_%"][::-1])): ax.text(p + 5e4, i, f"{m}% margin", va="center", fontsize=9)
ax.set_xlim(0, c.Profit.max() * 1.25); ax.set_title("Profit by Country"); save(fig, "04_country.png")

# 5 Product
fig, ax = plt.subplots(figsize=(7.5, 4.4)); p = prod.reset_index()
ax.bar(p.Product, p.Sales, color=NAVY, label="Sales"); ax.bar(p.Product, p.Profit, color=TEAL, label="Profit")
ax.yaxis.set_major_formatter(M); ax.set_title("Sales & Profit by Product"); ax.legend(frameon=False); save(fig, "05_product.png")

# 6 Heatmap segment x discount band
fig, ax = plt.subplots(figsize=(7, 4.4)); hm = seg_band_margin[order]
sns.heatmap(hm, annot=True, fmt=".1f", cmap="RdYlGn", center=10, cbar_kws={"label": "Margin %"}, ax=ax, linewidths=.5)
ax.set_title("Profit Margin % - Segment x Discount Band"); ax.set_xlabel(""); ax.set_ylabel(""); save(fig, "06_heatmap.png")

# 7 Like-for-like
fig, ax = plt.subplots(figsize=(5.5, 4.4)); ax.bar(["Sep-Dec 2013", "Sep-Dec 2014"], like.Sales, color=[GREY, NAVY])
ax.yaxis.set_major_formatter(M); ax.text(1, like.Sales.iloc[1] + 5e5, f"+{kpi['yoy_sales_pct']:.1f}%", ha="center", fontweight="bold", color=TEAL)
ax.set_title("Like-for-Like Growth (same 4 months)"); save(fig, "07_like_for_like.png")

# 8 Enterprise profit by discount band
fig, ax = plt.subplots(figsize=(7, 4.4)); e = ent.groupby("Discount_Band").Profit.sum().reindex(order)
ax.bar(order, e, color=[TEAL if v >= 0 else RED for v in e]); ax.axhline(0, color="k", lw=.6); ax.yaxis.set_major_formatter(M)
ax.set_title("Enterprise: Profit by Discount Band"); save(fig, "08_enterprise_loss.png")
print("Charts saved.")
