-- ============================================================
-- PROJECT: Healthcare Revenue Cycle Management (RCM) Analysis
-- Author:  Uma Maheshwari Nagarade
-- Tool:    SQL Server Management Studio (SSMS)
-- Dataset: Healthcare_RCM_Dataset_10k.xlsx (10,000 claims)
-- ============================================================

-- NOTE: After importing Excel into SQL Server,
-- The table name will be [Healthcare_RCM_Project].[dbo].[Claims_Data$]
-- Created a clean view first for easy querying.

-- ============================================================
-- STEP 1: CREATE A CLEAN VIEW
-- (So we don't have to type the long table name every time)
-- ============================================================

CREATE VIEW RCM_Claims AS
SELECT * FROM [HealthcareRCM].[dbo].[Claims_Data$];

-- Now you can simply use: SELECT * FROM RCM_Claims


-- ============================================================
-- STEP 2: EXPLORE THE DATA
-- (Always do this first — understand what you're working with)
-- ============================================================

-- Q: How many total records are there?
SELECT COUNT(*) AS Total_Claims
FROM RCM_Claims;

-- Q: What are the unique claim statuses?
SELECT DISTINCT Claim_Status
FROM RCM_Claims;

-- Q: What is the date range of the data?
SELECT 
    MIN(Claim_Date) AS Earliest_Claim,
    MAX(Claim_Date) AS Latest_Claim
FROM RCM_Claims;

-- Q: Preview first 10 rows
SELECT TOP 10 *
FROM RCM_Claims;


-- ============================================================
-- STEP 3: CLAIM STATUS ANALYSIS
-- (How many claims are Paid, Denied, Pending, Partially Paid?)
-- ============================================================

SELECT 
    Claim_Status,
    COUNT(*) AS Total_Claims,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER(), 2) AS Percentage
FROM RCM_Claims
GROUP BY Claim_Status
ORDER BY Total_Claims DESC;

--		WHAT THIS TELLS US:
-- Shows claim distribution — what % is paid vs denied vs pending
-- This is Our first KPI — Claim Status Breakdown


-- ============================================================
-- STEP 4: DENIAL RATE CALCULATION
-- (Key KPI — what % of claims are being denied?)
-- ============================================================

SELECT 
    COUNT(*) AS Total_Claims,
    SUM(CASE WHEN Claim_Status = 'Denied' THEN 1 ELSE 0 END) AS Denied_Claims,
    ROUND(
        SUM(CASE WHEN Claim_Status = 'Denied' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 
        2
    ) AS Denial_Rate_Percentage
FROM RCM_Claims;

--		WHAT THIS TELLS US:
-- Overall denial rate — benchmark is under 10% for healthy RCM
-- If it's above 20%, there's a serious problem


-- ============================================================
-- STEP 5: DENIAL RATE BY PAYER
-- (Which insurance companies deny the most?)
-- ============================================================

SELECT 
    Payer,
    COUNT(*) AS Total_Claims,
    SUM(CASE WHEN Claim_Status = 'Denied' THEN 1 ELSE 0 END) AS Denied_Claims,
    CAST(CAST(ROUND(
    SUM(CASE WHEN Claim_Status = 'Denied' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 
    2
    ) AS Denial_Rate_Pct
FROM RCM_Claims
GROUP BY Payer
ORDER BY Denial_Rate_Pct DESC;

--	   WHAT THIS TELLS US:
-- Which payers are most problematic
-- Helps teams focus follow-up efforts on high-denial payers


-- ============================================================
-- STEP 6: DENIAL REASONS ANALYSIS
-- (WHY are claims being denied? Root cause analysis)
-- ============================================================

SELECT 
    Denial_Reason,
    COUNT(*) AS Total_Denials,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER(), 2) AS Percentage
FROM RCM_Claims
WHERE Claim_Status IN ('Denied', 'Partially Paid')
  AND Denial_Reason IS NOT NULL
  AND Denial_Reason != ''
GROUP BY Denial_Reason
ORDER BY Total_Denials DESC;

--  WHAT THIS TELLS US:
-- Top root causes of denials
-- "Medical Necessity" and "Missing Information" are usually top 2
-- This directly feeds insight: "Identified top 3 denial reasons"


-- ============================================================
-- STEP 7: DENIAL RATE BY DEPARTMENT
-- (Which hospital department has highest denial rate?)
-- ============================================================

SELECT 
    Department,
    COUNT(*) AS Total_Claims,
    SUM(CASE WHEN Claim_Status = 'Denied' THEN 1 ELSE 0 END) AS Denied_Claims,
    ROUND(
        SUM(CASE WHEN Claim_Status = 'Denied' THEN 1 ELSE 0 END) * 100.0 / COUNT(*),
        2
    ) AS Denial_Rate_Pct
FROM RCM_Claims
GROUP BY Department
ORDER BY Denial_Rate_Pct DESC;


-- ============================================================
-- STEP 8: REVENUE ANALYSIS
-- (How much was billed vs actually collected?)
-- ============================================================

SELECT 
    SUM(Billed_Amount) AS Total_Billed,
    SUM(Paid_Amount) AS Total_Collected,
    SUM(Balance_Amount) AS Total_Outstanding,
    CAST(ROUND(SUM(Paid_Amount) * 100.0 / SUM(Billed_Amount), 2) AS varchar) + '%' AS Collection_Rate_Pct
FROM RCM_Claims;

--	  WHAT THIS TELLS US:
-- Collection rate — how efficiently revenue is being collected
-- Outstanding balance — how much is still unpaid


-- ============================================================
-- STEP 9: MONTHLY PAYMENT TRENDS
-- (How do billed and paid amounts trend over months?)
-- ============================================================

SELECT 
    FORMAT(CAST(Claim_Date AS DATE), 'yyyy-MM') AS Month,
    COUNT(*) AS Total_Claims,
    ROUND(SUM(Billed_Amount), 2) AS Total_Billed,
    ROUND(SUM(Paid_Amount), 2) AS Total_Paid,
    CAST(ROUND(SUM(Paid_Amount) * 100.0 / SUM(Billed_Amount), 2) AS varchar) + '%' AS Collection_Rate_Pct
FROM RCM_Claims
GROUP BY FORMAT(CAST(Claim_Date AS DATE), 'yyyy-MM')
ORDER BY Month;

--	  WHAT THIS TELLS US:
-- Monthly trends — great for a line chart in Power BI
-- Spot months where collection dropped — why did that happen?


-- ============================================================
-- STEP 10: AR AGING ANALYSIS
-- (How long are claims sitting unpaid?)
-- ============================================================

SELECT 
    AR_Bucket,
    COUNT(*) AS Total_Claims,
    ROUND(SUM(Balance_Amount), 2) AS Outstanding_Amount,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER(), 2) AS Percentage
FROM RCM_Claims
WHERE Claim_Status IN ('Pending', 'Partially Paid')
GROUP BY AR_Bucket
ORDER BY 
    CASE AR_Bucket 
        WHEN '0-30' THEN 1 
        WHEN '31-60' THEN 2 
        WHEN '61-90' THEN 3 
        WHEN '90+' THEN 4 
    END;

--    WHAT THIS TELLS US:
-- 90+ bucket is the most dangerous — revenue at risk of write-off
-- Focus collections effort on 90+ bucket first


-- ============================================================
-- STEP 11: PAYER PERFORMANCE
-- (Which payers pay fastest? Which are slowest?)
-- ============================================================

SELECT 
    Payer,
    COUNT(*) AS Total_Claims,
    ROUND(AVG(CAST(Days_to_Payment AS FLOAT)), 1) AS Avg_Days_to_Payment,
    ROUND(SUM(Paid_Amount), 2) AS Total_Paid,
    ROUND(SUM(Paid_Amount) * 100.0 / SUM(Billed_Amount), 2) AS Collection_Rate_Pct
FROM RCM_Claims
WHERE Claim_Status = 'Paid'
GROUP BY Payer
ORDER BY Avg_Days_to_Payment ASC;

--   WHAT THIS TELLS US:
-- Best and worst payers by speed and collection rate
-- Useful for payer contract negotiations


-- ============================================================
-- STEP 12: FINAL SUMMARY VIEW
-- (One query that gives all KPIs together)
-- ============================================================

SELECT 
    COUNT(*) AS Total_Claims,
    SUM(CASE WHEN Claim_Status = 'Paid' THEN 1 ELSE 0 END) AS Paid_Claims,
    SUM(CASE WHEN Claim_Status = 'Denied' THEN 1 ELSE 0 END) AS Denied_Claims,
    SUM(CASE WHEN Claim_Status = 'Pending' THEN 1 ELSE 0 END) AS Pending_Claims,
    SUM(CASE WHEN Claim_Status = 'Partially Paid' THEN 1 ELSE 0 END) AS Partially_Paid,
    ROUND(SUM(Billed_Amount), 2) AS Total_Billed,
    ROUND(SUM(Paid_Amount), 2) AS Total_Collected,
    ROUND(SUM(Balance_Amount), 2) AS Total_Outstanding,
    ROUND(SUM(CASE WHEN Claim_Status = 'Denied' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2) AS Denial_Rate_Pct,
    ROUND(SUM(Paid_Amount) * 100.0 / SUM(Billed_Amount), 2) AS Collection_Rate_Pct,
    ROUND(AVG(CAST(NULLIF(Days_to_Payment, 0) AS FLOAT)), 1) AS Avg_Days_to_Payment
FROM RCM_Claims;

-- ============================================================

