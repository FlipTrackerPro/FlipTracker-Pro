# Changelog

## v0.6.0 — Packaging Cost Audit Rounding Fix

- Fixed false Packaging audit errors for low-cost units such as tape measured by the foot.
- Cost Per Unit remains rounded to two decimal places.
- Calculation Audit now rounds the expected Cost Per Unit to two decimal places before comparing.
- Audit messages now show the expected rounded dollar value.
- Added **Repair Packaging Cost Audit** menu command.
