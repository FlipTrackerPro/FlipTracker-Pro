# Changelog

## v0.5.5 — Whole Packaging Quantities and Days Format Fix

- Complete Sale packaging quantity fields now use whole-number increments.
- Server rejects decimal packaging quantities instead of silently saving them.
- Sales packaging quantity columns are formatted as integers and validated as non-negative whole quantities.
- Days in Inventory is explicitly formatted as a whole number, correcting cells that inherited a date format.
- Existing Inventory rows are recalculated and reformatted during upgrade.
- Sales column formatting now uses header-name mappings rather than fixed positions.
