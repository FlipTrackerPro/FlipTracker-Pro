# FlipTracker Pro v0.4 Full Calculation Audit

## Result

The original v0.4.0 package did **not** pass the calculation audit. This v0.4.1 package contains the confirmed fixes.

## Confirmed defects fixed

1. **Critical — Sales profit understated when GST/HST was entered.** Tax collected was subtracted as a selling cost even though Gross Revenue already excluded it.
2. **Critical — Tax Centre year filters were fragile.** Full-column `YEAR()` expressions inside `FILTER()` could fail or calculate inconsistently. Tax calculations now use Apps Script date comparisons.
3. **Critical — Prior-year ending inventory was wrong after an item was sold in a later year.** Ending inventory now uses sale dates, not only the current inventory status.
4. **High — GST/HST ITCs were deducted even when the business was not registered or selected Quick Method.** ITCs now apply only when Registered = Yes and Reporting Method = Regular.
5. **High — Dashboard always deducted mileage.** Mileage is now deducted only when the Settings toggle is Yes.
6. **High — Inventory quantity did not affect Total Cost.** Total Cost is now Purchase Price × Quantity + Tax + Acquisition Shipping.
7. **High — Business kilometres could exceed total kilometres, and zero could be replaced by the total.** Values are now clamped and blank is distinguished from zero.
8. **Medium — Packaging quantity on hand defaulted to zero.** Blank now defaults to Units Purchased.
9. **High — Regular Method COGS and ending inventory included recoverable GST/HST while also claiming the ITC.** Tax-basis inventory values now exclude recoverable tax.
10. **Medium — Dashboard Inventory Cost included sold and archived inventory.** It now reports current inventory cost.
11. **Medium — Install/Repair erased customized Admin lookup entries.** Defaults are now merged with existing entries.

## Validation performed

- Header-to-column mapping review for all six transaction sheets
- Save-function arithmetic review
- Dashboard cross-sheet reference review
- Tax-year boundary and later-sale inventory tests
- GST/HST registered/unregistered logic tests
- Accountant Export link review
- Migration and installer non-destructive-data review
- Duplicate function and JavaScript syntax scan
- Numerical scenario reconciliation

## Remaining limitations

- FlipTracker Pro is bookkeeping software, not tax-filing software.
- The inventory row represents one complete item or lot. Recording a sale marks that whole row sold. Partial-lot sales are not supported yet.
- Users must avoid recording transaction-specific shipping, packaging, fees, or promotion costs again in Expenses, or reports will double count them.
- Quick Method GST/HST calculations are not estimated; net GST/HST displays zero unless Regular Method is selected.
- CRA mileage rates and tax treatment must be confirmed by the user or accountant.
