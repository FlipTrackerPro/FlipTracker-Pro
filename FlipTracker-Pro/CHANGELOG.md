# Changelog

## v0.4.9 — Header-Based Sales Save

- Replaced positional Sales writes with header-based record mapping.
- Complete Sale now writes Item ID and Description to their exact named columns.
- Added explicit checks for missing Sales columns and invalid sale dates.
- Corrected packaging dropdown target columns after Description was added.
- Added rollback if packaging or Inventory finalization fails.
- Added server-side error logging and a success response containing the Sales row.
