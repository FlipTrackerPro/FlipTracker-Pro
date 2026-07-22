# Changelog

## v0.5.1 — Complete Sale and Cost Fix

- Fixed Complete Sale form saving by exposing public Apps Script endpoints to HTML forms.
- Fixed the same private-function issue in Packaging, Expenses, and Mileage forms.
- Box Used now displays and stores only the Packaging ID.
- Shows box description, cost per unit, and available quantity separately beneath the Box Used field.
- Validates that the selected Box Packaging ID is actually categorized as Box.
- Packaging Cost is calculated from Packaging ID → Cost Per Unit × quantity.
- Added live Total Cost calculation to the Inventory form.
- Server-side Inventory Total Cost remains Purchase Price × Quantity + Tax Paid + Acquisition Shipping.
