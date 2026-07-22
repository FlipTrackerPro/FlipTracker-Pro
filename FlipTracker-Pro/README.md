# FlipTracker Pro v0.5.1

## Important fixes

The Complete Sale form now calls a public Apps Script function and records the sale. Box Used stores only the Packaging ID in the Sales sheet. The description and costing information are display-only details from the Packaging sheet.

The Inventory form now displays Total Cost live using:

`Purchase Price × Quantity + Tax Paid + Acquisition Shipping`

## Upgrade

Replace all Apps Script files, run `upgradeFlipTrackerPro()`, and reload the spreadsheet.
