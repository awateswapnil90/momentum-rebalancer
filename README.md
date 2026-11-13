# Momentum Rebalancer

Static page that reads this month's orders from the "Momentum Rebalancer" Google Sheet (published **Feed** tab, CSV), shows them for review, and sends them to a Zerodha Kite basket (Kite Publisher offsite order execution) when you press OK.

## One-time setup
1. In the sheet: File → Share → Publish to web → choose the **Feed** tab and **CSV** → Publish → copy the link.
2. Open the GitHub Pages link → **Setup**: paste the Kite Publisher **API key** and the Feed link. Both stay in your browser only. Never paste the API secret.
3. In the Kite developer console, set the app's Redirect URL to the GitHub Pages link.

## Each month
1. Import the new scanner CSV into the sheet's Scanner tab.
2. Open the page → **Review sells** → OK → place in Kite.
3. After the sells fill, reload → **Review buys** → OK → place in Kite.
4. Update Holdings and the cash cell (Settings B5) in the sheet.

If Kite rejects MARKET orders from the basket, switch to **LIMIT**: ref price ± buffer, rounded to a valid tick.

Not investment advice. Nothing is placed until you confirm in Kite.
