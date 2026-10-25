/* =====================================================================
   SALES PERFORMANCE ANALYSIS - SQL (MySQL 8.0)
   Dataset : sales_clean.csv  (700 rows, Sep 2013 - Dec 2014)
   Author  : <Your Name>  |  Beeskilled x Axcentra Internship - Week 4
   ===================================================================== */

-- ===== 1. DATABASE & TABLE SETUP =====
CREATE DATABASE IF NOT EXISTS sales_db;
USE sales_db;

DROP TABLE IF EXISTS sales;
CREATE TABLE sales (
    segment             VARCHAR(30),
    country             VARCHAR(40),
    product             VARCHAR(30),
    units_sold          DECIMAL(10,1),
    manufacturing_price INT,
    sale_price          INT,
    gross_sales         DECIMAL(14,2),
    discounts           DECIMAL(14,2),
    sales               DECIMAL(14,2),
    cogs                DECIMAL(14,2),
    profit              DECIMAL(14,2),
    sale_date           DATE,
    month_number        INT,
    month_name          VARCHAR(15),
    sales_year          INT,
    discount_band       VARCHAR(10)
);
/* Import: in MySQL Workbench -> right-click table `sales` -> Table Data Import Wizard
   -> choose 01_Data/sales_clean.csv -> map columns by name (extra CSV columns can be skipped). */

-- ===== ANALYSIS QUERIES =====

-- Q1. Executive KPIs
SELECT COUNT(*)                                   AS orders,
       ROUND(SUM(gross_sales),0)                  AS gross_sales,
       ROUND(SUM(discounts),0)                    AS discounts,
       ROUND(SUM(sales),0)                        AS net_sales,
       ROUND(SUM(profit),0)                       AS profit,
       ROUND(SUM(profit)/SUM(sales)*100,1)        AS margin_pct,
       ROUND(SUM(units_sold),0)                   AS units
FROM sales;

-- Q2. Segment performance + share of total profit (window function)
SELECT segment,
       ROUND(SUM(sales),0)                                   AS net_sales,
       ROUND(SUM(profit),0)                                  AS profit,
       ROUND(SUM(profit)/SUM(sales)*100,1)                   AS margin_pct,
       ROUND(SUM(profit)/SUM(SUM(profit)) OVER()*100,1)      AS share_of_total_profit_pct
FROM sales
GROUP BY segment
ORDER BY net_sales DESC;

-- Q3. Rank countries by profit
SELECT country,
       ROUND(SUM(sales),0)                    AS net_sales,
       ROUND(SUM(profit),0)                   AS profit,
       ROUND(SUM(profit)/SUM(sales)*100,1)    AS margin_pct,
       RANK() OVER (ORDER BY SUM(profit) DESC) AS profit_rank
FROM sales
GROUP BY country;

-- Q4. Product performance
SELECT product,
       ROUND(SUM(units_sold),0)               AS units,
       ROUND(SUM(sales),0)                    AS net_sales,
       ROUND(SUM(profit),0)                   AS profit,
       ROUND(SUM(profit)/SUM(sales)*100,1)    AS margin_pct
FROM sales
GROUP BY product
ORDER BY profit DESC;

-- Q5. Does discounting hurt profit?  (blank band in raw data = no discount)
SELECT COALESCE(NULLIF(discount_band,''),'No Discount')       AS discount_band,
       COUNT(*)                                        AS orders,
       ROUND(SUM(discounts)/SUM(gross_sales)*100,1)    AS avg_discount_pct,
       ROUND(SUM(profit),0)                            AS profit,
       ROUND(SUM(profit)/SUM(sales)*100,1)             AS margin_pct
FROM sales
GROUP BY COALESCE(NULLIF(discount_band,''),'No Discount')
ORDER BY avg_discount_pct;

-- Q6. Segment x Discount band margin matrix (conditional aggregation = pivot)
SELECT segment,
       ROUND(SUM(CASE WHEN COALESCE(NULLIF(discount_band,''),'No Discount') = 'No Discount' THEN profit END)/SUM(CASE WHEN COALESCE(NULLIF(discount_band,''),'No Discount') = 'No Discount' THEN sales END)*100,1) AS no_discount_pct,
       ROUND(SUM(CASE WHEN discount_band='Low'    THEN profit END)/SUM(CASE WHEN discount_band='Low'    THEN sales END)*100,1) AS low_pct,
       ROUND(SUM(CASE WHEN discount_band='Medium' THEN profit END)/SUM(CASE WHEN discount_band='Medium' THEN sales END)*100,1) AS medium_pct,
       ROUND(SUM(CASE WHEN discount_band='High'   THEN profit END)/SUM(CASE WHEN discount_band='High'   THEN sales END)*100,1) AS high_pct
FROM sales
GROUP BY segment;

-- Q7. Monthly trend with month-over-month growth (LAG)
WITH monthly AS (
    SELECT sales_year, month_number,
           SUM(sales) AS net_sales, SUM(profit) AS profit
    FROM sales
    GROUP BY sales_year, month_number
)
SELECT sales_year, month_number,
       ROUND(net_sales,0) AS net_sales,
       ROUND(profit,0)    AS profit,
       ROUND((net_sales - LAG(net_sales) OVER (ORDER BY sales_year, month_number))
             / LAG(net_sales) OVER (ORDER BY sales_year, month_number) * 100,1) AS mom_growth_pct
FROM monthly
ORDER BY sales_year, month_number;

-- Q8. Like-for-like growth: Sep-Dec 2013 vs Sep-Dec 2014
--     (2013 only has 4 months of data, so comparing full years would be misleading)
SELECT sales_year,
       ROUND(SUM(sales),0)  AS net_sales,
       ROUND(SUM(profit),0) AS profit
FROM sales
WHERE month_number >= 9
GROUP BY sales_year
ORDER BY sales_year;

-- Q9. Cumulative profit over time (running total)
WITH monthly AS (
    SELECT sales_year, month_number, SUM(profit) AS profit
    FROM sales GROUP BY sales_year, month_number
)
SELECT sales_year, month_number, ROUND(profit,0) AS profit,
       ROUND(SUM(profit) OVER (ORDER BY sales_year, month_number),0) AS cumulative_profit
FROM monthly
ORDER BY sales_year, month_number;

-- Q10. Loss-making orders: where do they come from?
SELECT segment, discount_band,
       COUNT(*)                AS loss_orders,
       ROUND(SUM(profit),0)    AS total_loss
FROM sales
WHERE profit < 0
GROUP BY segment, discount_band
ORDER BY total_loss;

-- Q11. Unit economics BEFORE any discount, by segment & list-price tier
--      (Government sells at 3 different price tiers, so price tier is part of the grouping)
SELECT segment, sale_price                                          AS list_price,
       ROUND(SUM(cogs)/SUM(units_sold),0)                           AS cogs_per_unit,
       ROUND((SUM(gross_sales)-SUM(cogs))/SUM(gross_sales)*100,1)   AS gross_margin_before_discount_pct,
       ROUND(SUM(discounts)/SUM(gross_sales)*100,1)                 AS avg_discount_pct
FROM sales
GROUP BY segment, sale_price
ORDER BY gross_margin_before_discount_pct;

-- Q12. Top 3 most profitable Segment-Product combinations per country (ROW_NUMBER)
WITH ranked AS (
    SELECT country, segment, product, SUM(profit) AS profit,
           ROW_NUMBER() OVER (PARTITION BY country ORDER BY SUM(profit) DESC) AS rn
    FROM sales
    GROUP BY country, segment, product
)
SELECT country, segment, product, ROUND(profit,0) AS profit
FROM ranked
WHERE rn <= 3
ORDER BY country, rn;

-- Q13. What-if: profit if Enterprise had received NO discounts (upper bound, assumes same volume)
SELECT ROUND(SUM(profit),0)                AS enterprise_profit_actual,
       ROUND(SUM(profit)+SUM(discounts),0) AS enterprise_profit_no_discount
FROM sales
WHERE segment = 'Enterprise';
