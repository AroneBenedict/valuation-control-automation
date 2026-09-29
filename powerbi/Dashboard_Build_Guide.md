# Power BI dashboard - build guide (about 30-45 minutes)

I cannot generate a `.pbix` file outside Power BI Desktop, so this guide gives the exact steps. Everything you need (data, model, DAX) is in this project.

## 1. Load and model
1. Open **Power BI Desktop** > **Get data > Text/CSV** > load `data/Clean_Positions.csv`, rename the table **Positions**.
2. Load `data/Thresholds.csv`, rename the table **Thresholds**.
3. In Power Query check types: `ValuationDate` = Date; `Quantity`, `BookPrice`, `IndependentPrice`, `TolerancePct` = Decimal number. Close & Apply.
4. **Model view**: relationship `Positions[AssetClass]` (many) to `Thresholds[AssetClass]` (one), single direction.
5. Paste the calculated columns, then the measures, from `DAX_Measures.dax` (**Modeling > New column / New measure**).

## 2. Check your numbers before designing (must match)
| Card | Expected |
|---|---|
| OK Count | 220 |
| Breach Count | 22 |
| Positions Tested | 242 |
| Breach Rate | 9.1% |
| No Price Count | 6 |
| Unmapped Count | 2 |
| Stale Date Count | 5 |

The same figures come out of the Excel Summary, the VBA Run_Log, and the Access `qry_ReconcileCounts`, which is the point of the project: three tools, one answer.

## 3. Page 1 - "Valuation Control Overview"
| Visual | Fields |
|---|---|
| 4 KPI cards | Positions Tested, Breach Count, Breach Rate (format 0.0%), Net MV Diff (Breaches) |
| Clustered bar | Axis: Positions[AssetClass]; Values: Breach Count; sort descending; data labels on |
| Column chart | Axis: Positions[Desk]; Values: Breach Rate |
| Donut | Legend: Positions[Status]; Values: Total Positions |
| Slicers | Positions[AssetClass], Positions[Desk] |

## 4. Page 2 - "Exceptions & Data Quality"
| Visual | Fields |
|---|---|
| Table (filter Status = BREACH) | PositionID, Desk, AssetClass, Instrument, Variance %, Tolerance %, MV Diff. Sort by MV Diff, use conditional-format data bars |
| Table (filter Status = NO INDEPENDENT PRICE or UNMAPPED ASSET CLASS) | PositionID, Desk, AssetClass, Status |
| Cards | No Price Count, Unmapped Count, Stale Date Count |

## 5. Finishing
- Theme colours: navy `#1F3864` headers, red for breaches, amber for data-quality items.
- Add a text box: "Synthetic data - illustrative tolerances".
- Save as `Valuation_Control_Dashboard.pbix`, take a screenshot of each page for your portfolio / GitHub README.
- To refresh: replace `Clean_Positions.csv` (the VBA `RunAll` macro re-exports it) and click **Refresh**.
