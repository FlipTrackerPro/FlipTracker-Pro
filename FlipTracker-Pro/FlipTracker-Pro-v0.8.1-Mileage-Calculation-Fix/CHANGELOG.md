# Changelog

## v0.8.1 – Mileage Calculation Fix

- Removed the Google Maps Distance column from Mileage.
- Total Kilometres now calculates automatically as End Odometer minus Start Odometer (columns I minus H).
- Added automatic recalculation when either odometer value is edited directly.
- Updated the mileage form, schema migration, number formatting, and calculation audit.


## v0.8.0 — Smart Shipping
- Added editable Shipping Settings sheet.
- Added Smart Shipping Assistant.
- Recommends the smallest active box that fits entered dimensions.
- Calculates packaging cost and configured shipping estimate.
- Records dimensions, weight, carrier, service, suggested packaging, estimated shipping, and shipping variance on Sales.
- Preserves v0.7.0 Smart Mileage and existing records through schema migration 8.0.
