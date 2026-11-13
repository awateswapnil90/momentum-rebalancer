# Momentum Rebalancer

- **index.html** (GitHub Pages): reads this month's orders from the Google Sheet's published **Feed** tab, shows them for review, and sends them to a Zerodha Kite basket (Kite Publisher) when you press OK. A **Sync fills to sheet** button opens the sheet's sync script.
- **apps-script/Code.gs**: Google Apps Script bound to the sheet. After a Kite login it reads today's orders, keeps only filled CNC orders sent from this page (Kite tag `kc:pub` / `MOMRB`), and updates Holdings, Trade Log (realized P&L), cash, History and the Dashboard.

## One-time setup
1. Sheet → Extensions → Apps Script → paste `apps-script/Code.gs` → Save → Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone) → Deploy → authorize.
2. Reload the sheet → **Kite** menu → *Save API key & secret* (free Kite Connect **Personal** app).
3. **Kite** menu → *Show links* → put the URL in the Personal app's Redirect URL (developers.kite.trade).
4. Sheet → File → Share → Publish to web → **Feed** tab, CSV → Publish.
5. This page → Setup: Publisher API key, Feed CSV link, Sync link. Stored only in your browser. Never put an API secret on this page.

## Each month
Import the new scanner CSV into the Scanner tab → page: send sells → Sync → Reload → send buys → Sync. Sync on the same day as the trades.

Not investment advice. Nothing is placed until you confirm in Kite. P&L is gross of charges.
