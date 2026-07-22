# Changelog

## v0.5.3 — Inventory Cost Recalculation Fix

- Fixed malformed row-level Inventory formulas.
- Replaced fragile formulas with header-based recalculation on edits and saves.
- Total Cost now recalculates from Purchase Price × Quantity + Tax Paid + Acquisition Shipping.
- Added Repair Inventory Calculations menu action.
- Upgrade forcibly clears stale validation and recalculates all existing Inventory rows.
