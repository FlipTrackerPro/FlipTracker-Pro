# FlipTracker Pro v0.4.5

## Reliable automatic sale workflow

Google Apps Script edit triggers cannot reliably display an HTML modal dialog.
For that reason, v0.4.5 uses a dependable two-step workflow:

1. On the Inventory sheet, change the Status to **Sold**.
2. FlipTracker Pro changes it to **Sale Pending**, highlights the row, and shows a reminder.
3. Keep the row selected and choose:
   **FlipTracker Pro → Complete Selected Pending Sale**
4. The Record Sale form opens with the Inventory item preselected.
5. Saving the form creates the Sales record, deducts packaging, calculates profit,
   and changes the Inventory status to **Sold**.

Run `upgradeFlipTrackerPro()` after replacing the Apps Script files.
