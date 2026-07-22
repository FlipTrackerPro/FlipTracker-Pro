const FTP3 = {
  APP_NAME: 'FlipTracker Pro',
  VERSION: '0.7.1',
  ROWS: 1000,
  SHEETS: {
    DASHBOARD:'Dashboard', INVENTORY:'Inventory', SALES:'Sales',
    EXPENSES:'Expenses', MILEAGE:'Mileage', PACKAGING:'Packaging',
    ADMIN:'Admin', SETTINGS:'Settings', TAX_CENTRE:'Tax Centre', ACCOUNTANT_EXPORT:'Accountant Export', PACKAGING_LISTS:'Packaging Lists', MILEAGE_LISTS:'Mileage Lists'
  },
  COLORS: {
    NAVY:'#1F4E78', BLUE:'#5B9BD5', LIGHT_BLUE:'#D9EAF7',
    GREEN:'#70AD47', LIGHT_GREEN:'#E2F0D9', GOLD:'#FFF2CC',
    ORANGE:'#FCE5CD', LIGHT_RED:'#FCE4D6', GRAY:'#F3F6F9',
    BORDER:'#B7C9D6', WHITE:'#FFFFFF', TEXT:'#1F2937'
  },
  INVENTORY_HEADERS: [
    'Item ID','Description','Category','Condition','SKU','Barcode','Quantity',
    'Purchase Date','Purchase Location','Purchase Price','Tax Paid',
    'Acquisition Shipping','Total Cost','Status','Marketplace','Listing Date',
    'Storage Location','Listed Price','Expected Sale Price','Days in Inventory',
    'Inventory Health','Projected Profit','Projected ROI %','Created At','Updated At','Notes',
    'Photo Link','Receipt Link'
  ],
  SALES_HEADERS: [
    'Sale ID','Item ID','Description','Sale Date','Marketplace','Sale Price',
    'Shipping Charged','Shipping Actual','Packaging Cost',
    'Marketplace Fees','Payment Fees','Promotion Expense',
    'GST/HST Collected','Item Cost','Gross Revenue','Total Selling Costs',
    'Net Proceeds','Realized Profit','Realized ROI %','Days to Sell',
    'Buyer','Tracking Number','Notes','Created At',
    'Box Used','Box Qty','Bubble Wrap Used','Bubble Wrap Qty',
    'Mailer Used','Mailer Qty','Tape Used','Tape Qty',
    'Other Packaging Used','Other Packaging Qty','Packaging Verified'
  ],
  EXPENSE_HEADERS: [
    'Expense ID','Date','Category','Vendor','Description','Subtotal',
    'GST/HST Paid','Total','Business Use %','Deductible Amount',
    'Payment Method','Receipt Link','Notes','Created At'
  ],
  MILEAGE_HEADERS: [
    'Trip ID','Date','Start Location','End Location','Business Purpose',
    'Round Trip','Vehicle','Start Odometer','End Odometer','Google Maps Distance',
    'Total Kilometres','Business Kilometres','CRA Rate','Claim Amount','Route Link','Notes','Created At'
  ],
  PACKAGING_HEADERS: [
    'Packaging ID','Category','Item Name','Size','Unit of Measure',
    'Units Purchased','Purchase Cost','Cost Per Unit','Quantity On Hand',
    'Reorder Level','Supplier','SKU / Barcode','Product Link','Active',
    'Notes','Updated At'
  ]
};