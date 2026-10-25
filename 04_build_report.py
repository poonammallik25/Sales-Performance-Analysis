"""Builds 04_Report/Final_Report.pdf and 04_Report/Stakeholder_Summary.pdf from the analysis outputs.
Run AFTER 01_clean_and_analyze.py:   python 02_Code/04_build_report.py"""
import json, os
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, KeepTogether, PageBreak, CondPageBreak

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CH = os.path.join(ROOT, "03_Dashboard", "charts"); OUT = os.path.join(ROOT, "04_Report"); os.makedirs(OUT, exist_ok=True)
k = json.load(open(os.path.join(ROOT, "03_Dashboard", "kpis.json")))
NAVY, TEAL, ORANGE, RED, LIGHT = (colors.HexColor(h) for h in ("#1F3A5F", "#2A9D8F", "#E9A23B", "#D1495B", "#EEF2F7"))

ss = getSampleStyleSheet()
H1 = ParagraphStyle("H1", parent=ss["Heading1"], fontName="Helvetica-Bold", fontSize=15, textColor=NAVY, spaceBefore=10, spaceAfter=6)
H2 = ParagraphStyle("H2", parent=ss["Heading2"], fontName="Helvetica-Bold", fontSize=11.5, textColor=TEAL, spaceBefore=8, spaceAfter=4)
B = ParagraphStyle("B", parent=ss["BodyText"], fontName="Helvetica", fontSize=9.5, leading=13.5, spaceAfter=4)
BL = ParagraphStyle("BL", parent=B, leftIndent=12, bulletIndent=2)
CELL = ParagraphStyle("CELL", parent=B, fontSize=8.5, leading=11, spaceAfter=0)
CELLB = ParagraphStyle("CELLB", parent=CELL, fontName="Helvetica-Bold", textColor=colors.white)
TITLE = ParagraphStyle("T", parent=ss["Title"], fontName="Helvetica-Bold", fontSize=22, textColor=NAVY, alignment=0, spaceAfter=2)
SUB = ParagraphStyle("S", parent=B, fontSize=10.5, textColor=colors.HexColor("#5B6B7C"))
CALL = ParagraphStyle("CALL", parent=B, fontSize=10.5, leading=15, textColor=colors.white)

def img(name, w=15.5):
    from PIL import Image as PI
    iw, ih = PI.open(os.path.join(CH, name)).size
    return Image(os.path.join(CH, name), width=w * cm, height=w * cm * ih / iw)
def bullets(items): return [Paragraph(t, BL, bulletText="•") for t in items]
def table(rows, widths, head=True):
    data = [[Paragraph(str(c), CELLB if (head and i == 0) else CELL) for c in r] for i, r in enumerate(rows)]
    t = Table(data, colWidths=[w * cm for w in widths], repeatRows=1)
    st = [("GRID", (0, 0), (-1, -1), .4, colors.HexColor("#C9D3DE")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]
    if head: st += [("BACKGROUND", (0, 0), (-1, 0), NAVY)]
    st += [("BACKGROUND", (0, i), (-1, i), LIGHT) for i in range(2, len(rows), 2)]
    t.setStyle(TableStyle(st)); return t
def callout(txt, color=NAVY):
    t = Table([[Paragraph(txt, CALL)]], colWidths=[17 * cm]); t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), color), ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10), ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)])); return t
def kpi_strip():
    vals = [(f"${k['total_sales']/1e6:.1f}M", "Net Sales"), (f"${k['total_profit']/1e6:.1f}M", "Profit"), (f"{k['margin_pct']:.1f}%", "Profit Margin"),
            (f"{k['units']/1e6:.2f}M", "Units Sold"), (f"${k['discounts']/1e6:.1f}M", "Discounts Given"), (f"{k['loss_rows']}", "Loss-making Orders")]
    top = [Paragraph(f'<font size="15" color="#1F3A5F"><b>{v}</b></font>', ParagraphStyle("c", alignment=1, leading=20)) for v, _ in vals]
    bot = [Paragraph(f'<font size="7.5" color="#5B6B7C">{l}</font>', ParagraphStyle("d", alignment=1)) for _, l in vals]
    t = Table([top, bot], colWidths=[2.83 * cm] * 6); t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), LIGHT), ("BOX", (0, 0), (-1, -1), .5, colors.HexColor("#C9D3DE")), ("TOPPADDING", (0, 0), (-1, 0), 8), ("BOTTOMPADDING", (0, 1), (-1, 1), 7)])); return t
def footer(c, d):
    c.saveState(); c.setFont("Helvetica", 7.5); c.setFillColor(colors.HexColor("#5B6B7C"))
    c.drawString(2 * cm, 1.1 * cm, "Sales Performance Analysis  |  Beeskilled x Axcentra Internship - Week 4 Capstone"); c.drawRightString(A4[0] - 2 * cm, 1.1 * cm, f"Page {d.page}")
    c.setStrokeColor(TEAL); c.setLineWidth(1.2); c.line(2 * cm, A4[1] - 1.2 * cm, A4[0] - 2 * cm, A4[1] - 1.2 * cm); c.restoreState()

# numbers used in text
ent_up = 976_714; ent_new = k["ent_profit"] + ent_up; hi_up = 2_249_336
swing = k["ent_profit"] + k["ent_discounts"]

# ============================ FINAL REPORT ============================
S = [Paragraph("Sales Performance Analysis", TITLE),
     Paragraph("End-to-End Case Study: Where do we make money, where do we lose it, and what should we do?", SUB),
     Paragraph("Data: 700 orders &middot; 5 segments &middot; 5 countries &middot; 6 products &middot; Sep 2013 - Dec 2014", SUB), Spacer(1, 8), kpi_strip(), Spacer(1, 10),
     callout("<b>Headline:</b> Sales are growing fast (+36.9% like-for-like), but <b>discounting is eroding margin</b> and the <b>Enterprise segment loses money</b> "
             f"(-${abs(k['ent_profit'])/1e6:.1f}M). Fixing discount discipline in Enterprise and High-discount deals is the single biggest profit lever."),
     Paragraph("Executive Summary", H1)]
S += bullets([
    f"<b>Overall:</b> ${k['total_sales']/1e6:.1f}M net sales, ${k['total_profit']/1e6:.1f}M profit, {k['margin_pct']:.1f}% margin. Like-for-like (Sep-Dec), sales grew <b>+{k['yoy_sales_pct']:.1f}%</b> and profit <b>+{k['yoy_profit_pct']:.1f}%</b> from 2013 to 2014.",
    "<b>Profit engine:</b> Government generates <b>67%</b> of all profit ($11.4M at a 21.7% margin). Small Business adds another 24.5%.",
    f"<b>Profit leak:</b> Enterprise sells ${k['ent_sales']/1e6:.1f}M but loses ${abs(k['ent_profit'])/1e3:,.0f}K. Its list price ($125) is only 4% above unit cost ($120), so any discount above 4% makes an order unprofitable. All 58 loss-making orders are Enterprise orders with Medium or High discounts.",
    "<b>Discount effect:</b> Margin falls from <b>21.9%</b> (no discount) to 17.9% (low), 14.4% (medium) and <b>9.1%</b> (high). The pattern holds in every segment.",
    "<b>Hidden opportunity:</b> Channel Partners earn a <b>73%</b> margin but are only 1.5% of sales; Midmarket earns 28% on 2% of sales.",
    "<b>Seasonality:</b> Sales peak in <b>October and December 2014</b> (about $12M each)."])
S += [Paragraph("1. Objective, Data & Method", H1),
      Paragraph("<b>Objective:</b> find what drives profit and loss in 16 months of sales and recommend actions. <b>Workflow:</b> Python (Pandas, Matplotlib, Seaborn) for cleaning/EDA &rarr; SQL (MySQL 8) for analysis &rarr; Excel and Power BI for dashboards &rarr; this report.", B),
      Paragraph("Data cleaning log", H2),
      table([["#", "Issue found", "Action taken"],
             ["1", "Column ' Sales' had a leading space", "Headers trimmed; spaces replaced by underscores"],
             ["2", "Discount Band blank in 53 rows", "All 53 rows have Discounts = 0, so they are no-discount sales: relabelled <b>No Discount</b> (not deleted)"],
             ["3", "36 rows have half-units (e.g. 4,492.5)", "Kept as reported; flagged to the data owner"],
             ["4", "Manufacturing Price does not reconcile to COGS / Units (only 164 of 700 rows match)", "COGS used as reported for profit; flagged as a data-governance issue"],
             ["5", "2013 contains only Sep-Dec; 2014 is a full year", "Growth measured <b>like-for-like (Sep-Dec only)</b>. A full-year comparison would wrongly show +249%"],
             ["6", "Integrity checks", "Gross - Discounts = Sales and Sales - COGS = Profit verified on all 700 rows; 0 duplicates; 0 nulls after cleaning"]],
            [0.8, 6.4, 9.8])]

S += [CondPageBreak(11 * cm), Paragraph("2. Findings", H1), Paragraph("2.1 Segment performance: Government earns the profit, Enterprise loses it", H2), img("01_segment.png", 13.5)]
S += bullets(["Government: $52.5M sales, $11.4M profit (21.7% margin) = 67% of total profit.",
              "Small Business: $42.4M sales but only 9.8% margin; the second-largest profit pool ($4.1M).",
              f"Enterprise: ${k['ent_sales']/1e6:.1f}M sales (3rd largest) but <b>-3.1% margin</b>.",
              "Channel Partners (73.1%) and Midmarket (27.7%) have the best margins but together are only 3.5% of sales."])
S += [CondPageBreak(8 * cm), Paragraph("2.2 Discounts: deeper discount, thinner margin", H2),
      Table([[img("03_discount_margin.png", 8.2), img("06_heatmap.png", 8.2)]], colWidths=[8.5 * cm, 8.5 * cm])]
S += bullets(["The average discount in the High band is 12.5% vs 2.5% in the Low band; margin drops from 17.9% to 9.1%.",
              "In the heatmap every segment gets worse as discounts deepen. Enterprise is the only one that turns negative (-3.0% Medium, -9.5% High).",
              "245 orders (35%) are High-discount, so this is not an edge case."])
S += [CondPageBreak(9 * cm), Paragraph("2.3 Enterprise deep-dive: a pricing problem, not a demand problem", H2)]
S += [Table([[img("08_enterprise_loss.png", 8.2),
              table([["Enterprise metric", "Value"], ["List price / unit", "$125"], ["Cost (COGS) / unit", "$120"], ["Gross margin before discount", "4.0%"],
                     ["Average discount given", "6.9%"], ["Actual profit", f"-${abs(k['ent_profit'])/1e3:,.0f}K"], ["Profit if no discounts*", f"+${swing/1e3:,.0f}K"]], [4.6, 3.4])]],
            colWidths=[8.5 * cm, 8.5 * cm], style=[("VALIGN", (0, 0), (-1, -1), "MIDDLE")])]
S += [Paragraph("*Upper bound: assumes the same volume would have sold without discounts. Enterprise discounts total $1.46M, which is more than the entire loss.", ParagraphStyle("n", parent=B, fontSize=8, textColor=colors.HexColor("#5B6B7C"))),
      CondPageBreak(8 * cm), Paragraph("2.4 Country & product: balanced, so they are not the lever", H2),
      Table([[img("04_country.png", 8.2), img("05_product.png", 8.2)]], colWidths=[8.5 * cm, 8.5 * cm])]
S += bullets(["Country margins range 12.0% (USA) to 15.7% (Germany); the USA has the highest sales but the lowest margin.",
              "Product margins range 12.6% (Velo) to 15.9% (Amarilla). Paseo is the volume leader (30% of units) at a 14.5% margin."])
S += [CondPageBreak(9 * cm), Paragraph("2.5 Time trend: strong growth and Q4 peaks", H2), img("02_monthly_trend.png", 13.5)]
S += bullets([f"Like-for-like Sep-Dec: sales ${k['like_2013_sales']/1e6:.1f}M &rarr; ${k['like_2014_sales']/1e6:.1f}M (<b>+{k['yoy_sales_pct']:.1f}%</b>); profit +{k['yoy_profit_pct']:.1f}%, so growth did not come at the cost of margin.",
              "Peak months: Oct 2014 ($12.4M) and Dec 2014 ($12.0M). Nov 2014 dipped to $5.4M, so demand is lumpy."])
S += [CondPageBreak(12 * cm), Paragraph("3. Recommendations", H1),
      table([["#", "Action", "Evidence", "Illustrative impact*"],
             ["1", "<b>Reprice or cap discounts in Enterprise</b> (max Low band, ~2.5%) or raise list price", "4% pre-discount margin; all 58 loss orders are Enterprise Medium/High", f"+$0.98M profit; Enterprise from -$0.61M to +${ent_new/1e3:,.0f}K"],
             ["2", "<b>Require approval for High-band discounts</b> (>10%) across all segments", "High band margin 9.1% vs 17.9% Low; 245 orders", "+$2.25M if High were capped at the Medium rate (7.2%)"],
             ["3", "<b>Grow Channel Partners and Midmarket</b> (partner programme, targeted sales)", "73% and 28% margins on only 3.5% of sales", "Every extra $1M of Channel sales adds about $0.73M profit"],
             ["4", "<b>Protect Government</b> (account management, renewals)", "67% of profit from 300 orders", "Risk reduction: a 10% Government dip costs about $1.1M"],
             ["5", "<b>Plan for Q4 peaks</b>; run promotions in low months instead of discounting at peak", "Oct and Dec at about $12M vs $5-6M in other months", "Better stock/staffing; fewer unnecessary discounts"],
             ["6", "<b>Fix data governance</b> (Manufacturing Price vs COGS, half-units)", "Only 164 of 700 rows reconcile", "More reliable margin reporting"]],
            [1.4, 5.4, 5.2, 5.0]),
      Paragraph("*Priorities 1 and 2 overlap (Enterprise High orders appear in both) and must not be added together. All scenarios assume unchanged sales volume; real volume may fall if discounts are removed, so pilot before rolling out.", ParagraphStyle("n2", parent=B, fontSize=8, textColor=colors.HexColor("#5B6B7C"))),
      Paragraph("4. Limitations & Next Steps", H1)]
S += bullets(["Only 16 months of data and a 4-month 2013, so seasonality is based on one full year.",
              "Profit uses COGS as reported; the Manufacturing Price column could not be reconciled.",
              "No customer-level or price-elasticity data, so discount scenarios are illustrative upper bounds.",
              "Next: test Enterprise pricing in one country for a quarter, build a discount-approval dashboard, and add customer IDs to the dataset for retention analysis."])
SimpleDocTemplate(os.path.join(OUT, "Final_Report.pdf"), pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=1.8 * cm, bottomMargin=1.8 * cm,
                  title="Sales Performance Analysis - Final Report").build(S, onFirstPage=footer, onLaterPages=footer)

# ============================ STAKEHOLDER SUMMARY (1 page) ============================
T = [Paragraph("Sales Performance: What the Data Tells Us", TITLE), Paragraph("Stakeholder summary | Sep 2013 - Dec 2014 | 700 orders, 5 segments, 5 countries", SUB), Spacer(1, 8), kpi_strip(), Spacer(1, 10),
     callout("<b>The message:</b> We are growing fast, but we are giving away margin through discounts, and one segment (Enterprise) is losing money. Better discount discipline could add roughly <b>$1-2M of profit</b> without winning a single new customer."),
     Paragraph("What we found", H1)]
T += bullets([f"<b>Growth is strong.</b> Sep-Dec sales rose +{k['yoy_sales_pct']:.1f}% and profit +{k['yoy_profit_pct']:.1f}% year over year.",
              "<b>Government carries us:</b> 67% of profit. Losing a few Government accounts would hurt far more than any other segment.",
              f"<b>Enterprise loses money</b> (-${abs(k['ent_profit'])/1e3:,.0f}K) because it is priced only 4% above cost, then discounted by about 7%.",
              "<b>Discounts hurt:</b> margin drops from 21.9% with no discount to 9.1% with a high discount, in every segment.",
              "<b>Small, profitable pockets exist:</b> Channel Partners (73% margin) and Midmarket (28%) are under-sold.",
              "<b>Sales peak in Q4</b> (Oct and Dec at about $12M)."])
T += [Paragraph("What we recommend", H1)]
T += bullets(["<b>1. Cap Enterprise discounts at about 2.5% or raise its price</b> &rarr; about +$1.0M profit.",
              "<b>2. Require approval for any discount above 10%</b> &rarr; up to +$2.2M (overlaps with #1).",
              "<b>3. Invest in Channel Partners and Midmarket</b>; protect Government accounts.",
              "<b>4. Pilot first:</b> test the pricing change in one country for a quarter, because volume may react."])
T += [Spacer(1, 6), Table([[img("01_segment.png", 8.2), img("03_discount_margin.png", 8.2)]], colWidths=[8.5 * cm, 8.5 * cm]),
      Paragraph("Scenarios assume unchanged volume and are illustrative. Full method, SQL and code are in the project folder.", ParagraphStyle("n3", parent=B, fontSize=8, textColor=colors.HexColor("#5B6B7C")))]
SimpleDocTemplate(os.path.join(OUT, "Stakeholder_Summary.pdf"), pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=1.8 * cm, bottomMargin=1.8 * cm,
                  title="Sales Performance - Stakeholder Summary").build(T, onFirstPage=footer, onLaterPages=footer)
print("Reports built")
