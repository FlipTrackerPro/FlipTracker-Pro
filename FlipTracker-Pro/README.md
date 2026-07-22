# FlipTracker Pro v0.5.0

## Inventory order

Item ID, Description, Category, Condition, SKU, Barcode, Quantity,
Purchase Date, Purchase Location, Purchase Price, Tax Paid,
Acquisition Shipping, Total Cost, Storage Location, Receipt Link,
Photo Link, Notes, Status, Marketplace, Listing Date, Listed Price,
Expected Sale Price, Days in Inventory, Projected Profit,
Projected ROI %, Created At, Updated At.

## Cost logic

**Total Cost = Purchase Price × Quantity + Tax Paid + Acquisition Shipping**

Projected Profit uses Listed Price when it is greater than zero. Otherwise it uses
Expected Sale Price. Projected ROI equals Projected Profit divided by Total Cost.

## Upgrade

Replace all Apps Script files, run `upgradeFlipTrackerPro()`, then reload the spreadsheet.
After upgrading, run **FlipTracker Pro → Run Calculation Audit**.
