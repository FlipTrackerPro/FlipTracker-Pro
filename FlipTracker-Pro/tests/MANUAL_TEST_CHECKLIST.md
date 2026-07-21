# Manual Test Checklist — v0.4

- [ ] Upgrade preserves Inventory, Sales, Expenses, Mileage, and Packaging rows
- [ ] Existing Settings values remain unchanged
- [ ] Tax Centre uses Settings tax year
- [ ] Gross revenue matches Sales gross revenue for selected year
- [ ] COGS matches sold-item cost for selected year
- [ ] Expense category summary populates
- [ ] GST/HST collected and potential ITCs populate
- [ ] Ending inventory excludes sold items
- [ ] Accountant Export refreshes
- [ ] Version reports schema 4 and v0.4.0

## v0.4.2 packaging
- [ ] Existing Packaging rows migrate without data loss
- [ ] Existing Sales rows migrate without data loss
- [ ] Add box, bubble wrap, mailer, tape, and other supply
- [ ] Sale form displays available stock and unit cost
- [ ] Packaging cost equals sum of unit cost × quantity
- [ ] Sale profit includes calculated packaging cost
- [ ] Packaging stock deducts exactly once
- [ ] Insufficient stock blocks the sale
- [ ] Low-stock row highlights correctly
- [ ] Dashboard packaging cards reconcile
