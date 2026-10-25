# Power BI Dashboard - Step-by-Step Build Guide
Build time: about 60-90 minutes. Data file: `01_Data/sales_clean.csv` (already cleaned, 700 rows).
Goal: a 2-page dashboard that tells the story **"Growth is strong, but discounts are eroding margin and Enterprise loses money."**

---
## Step 1 - Load the data
1. Power BI Desktop -> **Home -> Get data -> Text/CSV** -> pick `sales_clean.csv` -> **Transform Data**.
2. In Power Query check types: `Date` = Date, `Units_Sold`, `Gross_Sales`, `Discounts`, `Sales`, `COGS`, `Profit` = Decimal number, `Year`, `Month_Number` = Whole number, everything else = Text.
3. Rename the query to **`sales`**. Remove columns you will not use: `COGS_per_Unit`, `Manufacturing_Price` (optional). **Close & Apply**.
4. Select `Month_Name` -> **Column tools -> Sort by column -> Month_Number** (so months sort Jan-Dec, not alphabetically).
5. Select `Discount_Band` -> add a helper sort: create a column `Band_Order` (see DAX below) and sort `Discount_Band` by it, so bars show No Discount > Low > Medium > High.

## Step 2 - Create the DAX measures
Right-click the `sales` table -> **New measure** (one by one). Put them in a display folder called "KPIs".

```DAX
Net Sales        = SUM ( sales[Sales] )
Profit           = SUM ( sales[Profit] )
Profit Margin %  = DIVIDE ( [Profit], [Net Sales] )
Units Sold       = SUM ( sales[Units_Sold] )
Discounts        = SUM ( sales[Discounts] )
Avg Discount %   = DIVIDE ( [Discounts], SUM ( sales[Gross_Sales] ) )
Loss Orders      = CALCULATE ( COUNTROWS ( sales ), sales[Profit] < 0 )

Profit Share % =
DIVIDE ( [Profit], CALCULATE ( [Profit], ALLSELECTED ( sales ) ) )

-- Like-for-like growth (2013 only has Sep-Dec, so compare the same 4 months)
Sep-Dec Sales 2013 = CALCULATE ( [Net Sales], sales[Year] = 2013, sales[Month_Number] >= 9 )
Sep-Dec Sales 2014 = CALCULATE ( [Net Sales], sales[Year] = 2014, sales[Month_Number] >= 9 )
LFL Sales Growth % = DIVIDE ( [Sep-Dec Sales 2014] - [Sep-Dec Sales 2013], [Sep-Dec Sales 2013] )
```
New calculated **column** (Table view -> New column):
```DAX
Band_Order =
SWITCH ( sales[Discount_Band], "No Discount", 1, "Low", 2, "Medium", 3, "High", 4 )
```
**Check your numbers** (put the measures in cards with no filters - they must match):
| Measure | Expected |
|---|---|
| Net Sales | $118,726,350 (118.7M) |
| Profit | $16,893,702 (16.9M) |
| Profit Margin % | 14.2% |
| Units Sold | 1,125,806 |
| Discounts | $9,205,248 (9.2M) |
| Loss Orders | 58 |
| LFL Sales Growth % | +36.9% |

## Step 3 - Page 1: "Executive Overview"
Canvas: 16:9. Background `#FFFFFF`; title bar (rectangle, fill `#1F3A5F`) with the text **Sales Performance Dashboard | Sep 2013 - Dec 2014**.

```
+--------------------------------------------------------------+
| TITLE BAR                                                    |
+----+----+----+----+----+----+  +-----------+                 |
|Sales|Prof|Marg|Unit|Disc|Loss|  | SLICERS   |                 |
+----+----+----+----+----+----+  | Year      |                 |
| Segment: Sales vs Profit  | Monthly trend (line)  | Segment   |
| (clustered column)        |                       | Country   |
+---------------------------+-----------------------+ Product   |
| Margin % by Discount Band | Profit by Country     | Band      |
| (column)                  | (bar)                 +-----------+
+--------------------------------------------------------------+
```
| Visual | Fields | Format tips |
|---|---|---|
| 6 x **Card** | Net Sales, Profit, Profit Margin %, Units Sold, Discounts, Loss Orders | Callout size 28; display units Millions; Loss Orders in red `#D1495B` |
| **Clustered column** | X: `Segment`; Y: Net Sales, Profit | Sales `#1F3A5F`, Profit `#2A9D8F`; data labels on; title: *Government drives 67% of profit* |
| **Line chart** | X: `Year_Month`; Y: Net Sales, Profit | Markers on; title: *Sales peak in Oct and Dec 2014* |
| **Clustered column** | X: `Discount_Band` (sorted by Band_Order); Y: Profit Margin % | Data labels as %; color `#E9A23B`; title: *Deeper discounts = thinner margin* |
| **Clustered bar** | Y: `Country`; X: Profit | Sort descending; color `#1F3A5F` |
| **Slicers** (tile style) | `Year`, `Segment`, `Country`, `Product`, `Discount_Band` | Sync across both pages (View -> Sync slicers) |

## Step 4 - Page 2: "Where we lose money"
| Visual | Fields | Insight shown |
|---|---|---|
| **Matrix** | Rows: `Segment`; Columns: `Discount_Band`; Values: Profit Margin % | Conditional formatting -> Background color -> rules: `< 0` = `#F8C9D0` (red), `> 0.2` = `#CDEBE7` (green). Enterprise turns red at Medium/High |
| **Column chart** | X: `Discount_Band`; Y: Profit; Filter on this visual: Segment = Enterprise | Enterprise loses money at Medium (-$155K) and High (-$622K) discounts |
| **Card** x3 (filtered to Enterprise) | Net Sales, Profit, Profit Margin % | $19.6M, -$0.61M, -3.1% |
| **Table** | `Segment`, `Country`, `Product`, `Discount_Band`, Net Sales, Profit | Filter `Profit < 0`; sort ascending -> the 58 loss-making orders; add data bars on Profit |
| **Text box** | "Enterprise list price $125 vs cost $120 per unit = 4% margin before any discount" | The root cause, in plain words |
| **Text box** | Recommendation: cap Enterprise discounts at about 2.5%; approve discounts above 10% | Call to action |

## Step 5 - Storytelling polish (what impresses reviewers)
- **Insight titles**, not label titles ("Enterprise loses money at Medium/High discounts" instead of "Profit by Discount Band").
- Max 3-4 colors: navy `#1F3A5F`, teal `#2A9D8F`, orange `#E9A23B`, red `#D1495B` (red only for negatives/losses).
- Align everything to a grid (View -> Gridlines / Snap to grid); consistent fonts (Segoe UI).
- Add **Tooltips**: hover on a segment bar to see Margin % and Profit Share %.
- Rename pages ("Executive Overview", "Where we lose money") and add a **Bookmark / page navigator** (Insert -> Buttons -> Navigator).
- Remove chart clutter: turn off gridlines, keep axis labels short, no 3D, no pie charts for more than 4 slices.

## Step 6 - Export & submit
1. Save as `Sales_Dashboard.pbix` (File -> Save as).
2. **File -> Export -> Export to PDF** -> `PowerBI_Dashboard.pdf` (easy for reviewers who do not have Power BI).
3. Take 2 screenshots (one per page): `powerbi_page1.png`, `powerbi_page2.png` -> put them in `03_Dashboard/` (and use them in the README).
4. Upload `.pbix`, `.pdf` and screenshots to the Google Drive folder `03_Dashboard/`.
