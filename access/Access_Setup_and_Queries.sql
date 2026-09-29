-- =============================================================================
-- Position Price Verification - MS Access layer   (SYNTHETIC DATA)
-- Access runs ONE statement per query, so create each block below as its own object:
--   Create > Query Design > close the table dialog > SQL View > paste one block > Save with the name shown.
-- =============================================================================

-- ---------- STEP 1: import the data -------------------------------------------
-- External Data > New Data Source > From File > Text File > "Import the source data into a new table".
--   Clean_Positions.csv  -> table  tbl_Positions   (set PositionID as Primary Key,
--                                                    ValuationDate = Date/Time, Quantity = Double,
--                                                    BookPrice / IndependentPrice = Double)
--   Thresholds.csv       -> table  tbl_Thresholds  (AssetClass = Primary Key, TolerancePct = Double)

-- ---------- STEP 2: save each block as a query ----------------------------------

-- Query name: qry_PriceCheck   (core control test: same rules as the Excel workbook)
SELECT p.PositionID, p.ValuationDate, p.Desk, p.AssetClass, p.Instrument, p.Currency,
       p.Quantity, p.BookPrice, p.IndependentPrice,
       IIf(p.IndependentPrice Is Null, Null, (p.BookPrice - p.IndependentPrice) / p.IndependentPrice) AS VariancePct,
       t.TolerancePct,
       IIf(p.IndependentPrice Is Null, Null, p.Quantity * (p.BookPrice - p.IndependentPrice)) AS MVDiff,
       IIf(t.TolerancePct Is Null, "UNMAPPED ASSET CLASS",
         IIf(p.IndependentPrice Is Null, "NO INDEPENDENT PRICE",
           IIf(Abs((p.BookPrice - p.IndependentPrice) / p.IndependentPrice) > t.TolerancePct, "BREACH", "OK"))) AS Status
FROM tbl_Positions AS p LEFT JOIN tbl_Thresholds AS t ON p.AssetClass = t.AssetClass;

-- Query name: qry_Exceptions   (breaches, largest market-value impact first)
SELECT PositionID, Desk, AssetClass, Instrument, Currency, Quantity, BookPrice, IndependentPrice, VariancePct, TolerancePct, MVDiff
FROM qry_PriceCheck
WHERE Status = "BREACH"
ORDER BY Abs(MVDiff) DESC;

-- Query name: qry_SummaryByAssetClass
SELECT AssetClass,
       Sum(IIf(Status = "OK" Or Status = "BREACH", 1, 0)) AS PositionsTested,
       Sum(IIf(Status = "OK", 1, 0)) AS OKCount,
       Sum(IIf(Status = "BREACH", 1, 0)) AS BreachCount,
       Sum(IIf(Status = "NO INDEPENDENT PRICE", 1, 0)) AS NoPriceCount,
       Sum(IIf(Status = "BREACH", 1, 0)) / Sum(IIf(Status = "OK" Or Status = "BREACH", 1, 0)) AS BreachRate,
       Sum(MVDiff) AS NetMVDiff
FROM qry_PriceCheck
WHERE Status <> "UNMAPPED ASSET CLASS"
GROUP BY AssetClass;

-- Query name: qry_SummaryByDesk
SELECT Desk,
       Sum(IIf(Status = "OK" Or Status = "BREACH", 1, 0)) AS PositionsTested,
       Sum(IIf(Status = "OK", 1, 0)) AS OKCount,
       Sum(IIf(Status = "BREACH", 1, 0)) AS BreachCount,
       Sum(IIf(Status = "NO INDEPENDENT PRICE", 1, 0)) AS NoPriceCount,
       Sum(IIf(Status = "BREACH", 1, 0)) / Sum(IIf(Status = "OK" Or Status = "BREACH", 1, 0)) AS BreachRate,
       Sum(MVDiff) AS NetMVDiff
FROM qry_PriceCheck
WHERE Status <> "UNMAPPED ASSET CLASS"
GROUP BY Desk;

-- Query name: qry_DataQuality   (unmapped labels and missing prices needing follow-up)
SELECT PositionID, Desk, AssetClass, Instrument, Status
FROM qry_PriceCheck
WHERE Status = "UNMAPPED ASSET CLASS" OR Status = "NO INDEPENDENT PRICE"
ORDER BY Status, PositionID;

-- Query name: qry_StaleDates   (Access prompts for the report date when opened, e.g. 2026-09-28)
PARAMETERS [ReportDate] DateTime;
SELECT PositionID, Desk, AssetClass, ValuationDate
FROM tbl_Positions
WHERE ValuationDate <> [ReportDate]
ORDER BY ValuationDate, PositionID;

-- Query name: qry_ReconcileCounts   (compare with Excel Summary / VBA Run_Log: 220 OK, 22 BREACH, 6 NO PRICE, 2 UNMAPPED)
SELECT Status, Count(*) AS Positions
FROM qry_PriceCheck
GROUP BY Status;
