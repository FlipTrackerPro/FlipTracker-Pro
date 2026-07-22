# FlipTracker Pro v0.5.4

## Updated Inventory order

Item ID, Description, Category, Condition, SKU, Barcode, Quantity,
Purchase Date, Purchase Location, Purchase Price, Tax Paid,
Acquisition Shipping, Total Cost, Status, Marketplace, Listing Date,
Storage Location, Listed Price, Expected Sale Price, Days in Inventory,
Projected Profit, Projected ROI %, Created At, Updated At, Notes,
Photo Link, Receipt Link.

## Formula logic

Total Cost = Purchase Price × Quantity + Tax Paid + Acquisition Shipping.

All active Inventory-dependent calculations use header-name mappings rather than fixed column numbers.

## Upgrade

Replace the Apps Script files, run `upgradeFlipTrackerPro()`, reload the spreadsheet,
and run **FlipTracker Pro → Repair Inventory Calculations**.
