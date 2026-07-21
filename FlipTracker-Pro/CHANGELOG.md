# Changelog

## v0.4.2 — Packaging Integration
- Expanded Packaging into a master supply inventory.
- Added packaging IDs, categories, item names, units of measure, supplier SKUs, and active status.
- Added sales fields for box, bubble wrap, mailer, tape, and other packaging usage.
- Automatically calculates packaging cost and includes it in realized profit.
- Automatically deducts packaging stock with negative-stock protection.
- Added low-stock highlighting and dashboard packaging metrics.
- Added schema 4.2 migration preserving existing Sales and Packaging rows.
- Preserved all v0.4.1 audited calculation fixes.

# Changelog

## v0.4.1 — Calculation Audit Fixes
- Corrected GST/HST treatment in realized profit.
- Rebuilt Tax Centre calculations using reliable date-based Apps Script logic.
- Corrected historical ending inventory using actual sale dates.
- Added registration/method-aware ITC calculations.
- Corrected quantity, mileage, packaging, and dashboard calculations.
- Added an in-workbook Calculation Audit command and detailed audit report.

# Changelog

## v0.4.0
- Added CRA Tax Centre with selectable tax year.
- Added gross revenue, COGS, gross profit, expenses, and net-income estimates.
- Added GST/HST regular-method estimate and potential ITC summaries.
- Added year-end inventory count and value.
- Added accountant-ready export sheet.
- Added schema 4 migration that preserves existing Settings values.
- Added official CRA reference links in the Tax Centre.

## v0.3.0
- Added sales, expenses, mileage, packaging, dashboard, and structured repository layout.
