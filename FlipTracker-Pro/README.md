# FlipTracker Pro v0.5.5

This release fixes Complete Sale packaging quantities and the Inventory Days in Inventory display.

## Packaging quantities

Box, bubble wrap, mailer, tape, and other packaging quantities accept whole numbers only. The Sales sheet stores and displays them with no decimal places.

## Days in Inventory

Days in Inventory is stored and displayed as an integer, not as a date.

## Upgrade

Replace all Apps Script files, run `upgradeFlipTrackerPro()`, and reload the spreadsheet.
