# Changelog

## v0.5.0 — Inventory Flow and Formula Audit

- Reordered Inventory columns to follow identification → purchase → documentation → listing → performance → system fields.
- Kept the column name **Purchase Price** unchanged.
- Total Cost remains: Purchase Price × Quantity + Tax Paid + Acquisition Shipping.
- Projected Profit now uses Listed Price when present, otherwise Expected Sale Price.
- Automatically fills Listing Date when Status first changes to Listed.
- Hid Created At and Updated At from normal Inventory entry.
- Converted Inventory-dependent Sales, Dashboard, Tax Centre, and Audit logic to header-based lookups.
- Added formula auditing for Total Cost, Projected Profit, Projected ROI, sales totals, packaging costing, expenses, and mileage.
- Preserved existing data through header-name migration.
