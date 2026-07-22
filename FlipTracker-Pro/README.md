# FlipTracker Pro v0.4.7

## Fixed sale saving

The sale is now written and finalized without rebuilding the full Dashboard while
the document lock is held. This prevents the form remaining stuck on **Saving…**.

If a response takes longer than 20 seconds, the form displays a warning. Check the
Sales sheet before retrying because a slow connection may have completed the save.

## Marketplace dropdown

Marketplace values are taken from the **Marketplaces** column on the Admin sheet.
Edit that list to control the choices displayed in the Record Sale form.

## Upgrade

Replace the Apps Script files, run `upgradeFlipTrackerPro()`, and reload the spreadsheet.
