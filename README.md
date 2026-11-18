# Swiggy & Instamart Order Exporter

Exports your own Swiggy Food or Instamart order history to Excel for a date range. One tab per month plus a Summary tab.

## Install (Brave or Chrome)
1. Unzip this folder somewhere permanent.
2. Open `brave://extensions` (or `chrome://extensions`) and turn on Developer mode.
3. Click "Load unpacked" and pick the unzipped `swiggy-order-exporter` folder.
   (If you installed v1.0, click the reload icon on the extension after replacing the files.)

## Use
1. Log in and keep a https://www.swiggy.com tab (food) or https://instamart.in tab (Instamart) active.
2. Click the extension icon. It detects which site you are on.
3. Pick start and end dates (or Last month / This month / Last 3 months).
4. Leave "Delivered orders only" ticked to skip cancelled or failed orders.
5. On Instamart, "Include bill details" also reads each order's bill (fees, discounts). Untick it for a faster, lighter export.
6. Click "Export to Excel". A progress box appears on the page; the .xlsx downloads when it finishes.

Run it once on each site. Files are named `swiggy-orders_<start>_to_<end>.xlsx` and `instamart-orders_<start>_to_<end>.xlsx`.

## Columns
Swiggy: Order ID, Date, Time, Restaurant, Area, City, Items, Qty, Item total, Discount, Coupon, Delivery fee, Packing, Platform fee, Taxes, Tip, Amount paid, Payment, Status.

Instamart: Order ID, Date, Time, Area, Items, Qty, Item bill, Delivery fee, Handling fee, Other fees (e.g. surge or small-cart fees), Discount, Amount paid, Payment, Status.

## Notes
- It reads the same internal endpoints the sites' own order pages use, newest first, and stops once it passes your start date. Nothing leaves your browser.
- Instamart dates are converted to India time (IST).
- Amount paid is what you were charged, after discounts.
- These are internal endpoints and may change without notice; the field names in `exporter.js` are the place to fix if exports start failing.
