# Changelog

## v0.4.6 — Sales Form Save Fix

- Replaced the fragile submit handler with a direct **Accept Sale** action.
- Added client-side validation for required sale fields and packaging selections.
- Added a visible Saving state and disabled duplicate clicks while saving.
- Added clear success and server-error messages inside the form.
- Escaped dropdown labels to prevent special characters from breaking the HTML form.
- The server now returns the saved Sale ID after a successful save.
