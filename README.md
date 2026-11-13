# Momentum Rebalancer

Static page that sends an order list from the "Momentum Rebalancer" Google Sheet to a Zerodha Kite basket (Kite Publisher offsite order execution).

1. Open the GitHub Pages link.
2. Paste your Kite Publisher **API key** (stored only in your browser; never paste the API secret).
3. Paste the JSON from the sheet's **Orders** tab (sell list first, then buy list).
4. Check the preview, tap **Send to Kite basket**, review in Kite and place.

If Kite rejects MARKET orders from the basket, switch to **LIMIT**: ref price ± buffer, rounded to a valid tick.

Not investment advice. Nothing is placed until you confirm in Kite.
