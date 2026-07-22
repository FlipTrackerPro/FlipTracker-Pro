# FlipTracker Pro v0.4.8

## Corrected sale data model

Inventory and Sales now use matching fields:

- **Item ID**
- **Description**

On the Complete Sale form, the Item ID dropdown contains only Inventory Item IDs.
The Description appears in its own read-only field and is copied from Inventory when
the sale is saved.

## Upgrade

Replace the Apps Script files, run `upgradeFlipTrackerPro()`, and reload the spreadsheet.

The upgrade preserves existing Inventory data by moving **Title** into **Description**.
It also fills the new Sales Description column by matching existing Sales Item IDs to Inventory.
