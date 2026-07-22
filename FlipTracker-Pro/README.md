# FlipTracker Pro v0.6.0

This release fixes false Cost Per Unit audit errors.

Example:

- Purchase Cost: `$2.99`
- Units Purchased: `125 feet`
- Exact calculation: `$0.02392` per foot
- Stored two-decimal Cost Per Unit: `$0.02`

The Calculation Audit now compares `$0.02` with `$0.02`, instead of comparing the
stored rounded value with the unrounded calculation.

## Upgrade

1. Replace all Apps Script files.
2. Run `upgradeFlipTrackerPro()`.
3. Reload the spreadsheet.
4. Run **FlipTracker Pro → Repair Packaging Cost Audit**.
5. Run **FlipTracker Pro → Run Calculation Audit**.
