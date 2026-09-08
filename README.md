# Advanced-Excel---Retail-Analytics
Practiced raw data cleaning (Power Query) through pivot dashboards, cross-sheet formulas, and a VBA automation macro that generates region-filtered PDF reports on demand.

Phase 1 - Foundations: formatting, tables, simple formulas
☐  Convert Orders_RAW to an Excel Table (Ctrl+T), name it tbl_Orders.
☐  Apply currency format to UnitPrice, percentage format to Discount, a proper date format to OrderDate.
☐  Add a calculated column LineTotal = Quantity * UnitPrice * (1 - Discount).
☐  Convert Customers, Products, Employees, Returns, Regional_Targets, Region_Lookup into named Tables.
☐  On a new Summary sheet, report total revenue, average order value, order count, largest and smallest order using SUM / AVERAGE / COUNT / MAX / MIN.
☐  Conditional-format the top 10% of LineTotal values, and flag blank Quantity rows.

Phase 2 - Lookups and logic
☐  Add CustomerName to Orders_RAW via XLOOKUP (or VLOOKUP), pulling from Customers by CustomerID.
☐  Redo the same lookup with INDEX/MATCH in a second helper column and confirm both match.
☐  Add ProductName, Category, and Cost from Products the same way.
☐  Add Margin = LineTotal - (Quantity * Cost).
☐  Add MarginBand using nested IF (or IFS): High >30%, Medium 15-30%, Low otherwise.
☐  Add IsReturned - Yes/No flag checking whether OrderID appears in Returns (MATCH/ISNUMBER or COUNTIF).
☐  Add RegionName via lookup against Region_Lookup - note where messy Region codes break the match.

Phase 3 - Pivot tables and charts
☐  Pivot1: Revenue by Region by Quarter (add a helper Quarter column first).
☐  Pivot2: Revenue and Margin % by Category, sorted descending by revenue.
☐  Pivot3: Top 10 Customers by Revenue.
☐  Pivot4: Return Rate by Category (average of the IsReturned 1/0 flag).
☐  Pivot5: Sales by Employee, filtered to a single region using a slicer.
☐  Chart1: clustered column - Revenue by Region by Quarter.
☐  Chart2: pie or bar - Revenue by Category.
☐  Chart3: combo (columns + line) - Revenue vs Return Rate by Category.
☐  Create named ranges CurrentFXRate and ReportingYear and use each in a live formula.
☐  On Regional_Targets, add ActualSales via a two-criteria SUMIFS (Region + Quarter) and a VarianceToTarget column.
☐  Build a small INDEX/MATCH lookup keyed on a concatenated Region&Quarter for a typed-in Region+Quarter.

Phase 4 - Complex modelling, array/dynamic formulas
☐  Commission model: 3% on the first R50,000 of an employee's quarterly sales, 5% on the next R100,000, 7% above that - build with nested IF, then again with a single SUMPRODUCT-based formula, and confirm they match.
☐  If available: UNIQUE(Orders_Clean[Category]) to spill a distinct category list.
☐  If available: SORT(FILTER(...)) to spill all order lines above R5,000, sorted.
☐  If available: a SEQUENCE-generated 12-month calendar spine so zero-sales months show as 0, not blank.
☐  Build a Data Table (What-If Analysis) showing total margin across discount levels (0/5/10/15/20%) and cost-inflation assumptions (0/5/10%).
☐  Compute total revenue for Category = Technology AND MarginBand = High in one cell using SUMPRODUCT (no helper columns, no pivot).
Deliverable: a Model sheet with the commission engine, at least one dynamic array formula, and a scenario data table.

Phase 5 - Automation (VBA or Office Scripts)
☐  Choose one path: VBA (desktop Excel) or Office Scripts (Excel Online / Microsoft 365 web).
☐  Write a macro/script that refreshes the Power Query connection and all pivot tables.
☐  Export the Summary sheet as a dated PDF (e.g. Summary_2026-07-20.pdf).
☐  Add a Form Control button (VBA) or a documented run-manually step (Office Scripts) to trigger it.
☐  Stretch: parameterise it - prompt for a region and filter a pivot/slicer to it before exporting.
