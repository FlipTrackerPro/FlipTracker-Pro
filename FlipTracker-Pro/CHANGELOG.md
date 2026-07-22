# Changelog

## v0.4.5 — Reliable Sale Pending Workflow

- Fixed the v0.4.4 attempt to open an HTML form directly from `onEdit`.
- Selecting `Sold` now reliably changes the item to `Sale Pending`.
- Added a clear toast and cell note explaining how to complete the sale.
- Added **Complete Selected Pending Sale** to the FlipTracker Pro menu.
- The command opens a prefilled Record Sale form for the selected Inventory row.
- Saving clears the pending note and finalizes the item as `Sold`.
- Cancelling restores the previous status and clears the pending note.
- Duplicate Sales records remain blocked.
