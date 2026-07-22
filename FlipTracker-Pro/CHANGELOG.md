# Changelog

## v0.5.4 — Inventory Column Order Audit

- Moved Storage Location directly after Listing Date.
- Moved Notes, Photo Link, and Receipt Link to the end of Inventory in that order.
- Preserved existing Inventory data by migrating values by header name.
- Re-ran Inventory calculation repair after migration.
- Confirmed Total Cost, projected values, Sales item cost, Dashboard, Tax Centre, and Audit use header-name lookups.
- Replaced an older Description migration that still relied on fixed Inventory column positions.
