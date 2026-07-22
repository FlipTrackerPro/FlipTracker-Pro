# Changelog

## v0.4.8 — Item ID and Description Alignment

- Renamed Inventory **Title** to **Description**.
- Added **Description** as a separate Sales column immediately after Item ID.
- Changed the Complete Sale form Inventory Item selector to display only Item ID.
- Added a separate read-only Description field to the Complete Sale form.
- Sale saving now writes Item ID and Description into matching Sales columns.
- Existing Inventory Title data is migrated into Description.
- Existing Sales descriptions are populated by matching Item ID to Inventory.
- Updated Sales calculations, audits, tax reporting, dashboard formulas, packaging references, formats, and validations for the new column.
