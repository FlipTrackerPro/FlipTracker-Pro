# Sprint 3 — Transactions

Sprint 3 upgrades FlipTracker Pro from an inventory tracker into a transaction-based
resale business system.

## Added

- Sales entry form linked to active inventory
- Shipping Charged and Shipping Actual tracked separately
- Packaging cost per sale
- Marketplace, payment, and promotion fees
- GST/HST collected field
- Automatic gross revenue, selling costs, net proceeds, realized profit, ROI, and days to sell
- Automatic inventory status change to Sold
- Expense entry form with business-use percentage and deductible amount
- Mileage form with odometer calculation and CRA-rate claim
- Packaging supply catalogue with cost per unit and low-stock alerts
- Expanded dashboard with realized results

## Install

Use either:

- `src/` — organized Apps Script source files, or
- `dist/FlipTrackerPro_Sprint3.gs` — one deployable file.

In Google Sheets:

1. Open **Extensions → Apps Script**.
2. Back up your current Sprint 2 script.
3. Replace it with the Sprint 3 distribution file, or create one script file for each source module.
4. Save.
5. Run `initializeFlipTrackerProSprint3`.
6. Approve permissions when requested.
7. Refresh the spreadsheet.

Test on a copy of the workbook first.
