# Sprint 2 — Inventory Engine

## Recommended Apps Script installation

Create one Apps Script file for each file in `src/`, preserving the numeric order.
Google Apps Script loads all `.gs` files as one project, so the numeric prefixes are
for organization only.

Alternatively, paste the tested single-file build from:

`dist/FlipTrackerPro_Sprint2.gs`

## Features

- Add and edit inventory forms
- Stable sequential Item IDs
- Search by Item ID, title, SKU, and barcode
- Inventory aging alerts
- Slow-inventory filter
- Automatic cost, projected profit, ROI, and days-held calculations
- Created and updated timestamps
- Dashboard KPI improvements
- Admin-managed lookup lists

## First run

1. Open the FlipTracker Pro Google Sheet.
2. Select Extensions → Apps Script.
3. Add the source files or use the distribution build.
4. Save.
5. Run `initializeFlipTrackerProSprint2`.
6. Approve Google permissions.
7. Refresh the spreadsheet.
