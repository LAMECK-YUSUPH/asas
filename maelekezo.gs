# Setup guide

You're wiring together three things: a Google Sheet (storage), a Google Apps
Script (backend API), and `index.html` (the app you actually use).

## 1. Create the Sheet
1. Go to sheets.new — this becomes your permanent storage.
2. Name it something like "Shop Data".

## 2. Add the backend
1. In the Sheet, go to **Extensions → Apps Script**.
2. Delete anything in the editor and paste in the full contents of `Code.gs`.
3. Save (Ctrl/Cmd+S). Name the project e.g. "Shop Backend".
4. In the function dropdown at the top, select **setup**, then click **Run**.
   - First time, it'll ask you to authorize — approve it (it's your own script
     acting on your own sheet).
   - This creates three tabs: `Products`, `StockIn`, `Sales`, and pre-fills
     `Products` with your 15 items at zero stock/zero price. Use the
     **Register stock** page in the app to fill in real quantities and prices —
     don't type prices directly into the sheet, so the app's totals stay correct.

## 3. Deploy it as a Web App
1. Click **Deploy → New deployment**.
2. Type: **Web app**.
3. Execute as: **Me**.
4. Who has access: **Anyone** (this makes the URL your app's password — don't
   share it publicly; anyone with the link can call the API).
5. Click **Deploy**, authorize again if asked, and copy the **Web app URL**
   (ends in `/exec`).

## 4. Connect the frontend
1. Open `index.html` in a text editor.
2. Find the line near the top of the `<script>`:
   ```js
   const WEB_APP_URL = 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
   ```
3. Replace the placeholder with the URL you copied. Save.
4. Open `index.html` in your browser (double-click it, or host it anywhere —
   GitHub Pages, Netlify, or even a folder on your phone). That's the whole app.

## 5. Day to day
- Every time you buy stock from the main company → **Register stock**.
- Every sale → **Sell**, choose Wholesale or Retail, add items, **Complete sale**.
  Stock is deducted automatically and safely even if two staff sell at the same time.
- **Home** and **Reports** update live from the Sheet — nothing to refresh manually.

## If you change `Code.gs` later
Every time you edit the script, you must **Deploy → Manage deployments → edit
(pencil) → New version** for the changes to actually go live. Editing the code
alone does not update the running Web App.

## Known limits worth knowing
- **Apps Script quotas**: ~20,000 URL fetch calls/day free tier, and roughly
  300 requests/minute to Sheets — comfortably enough for a shop doing well
  under a thousand sales a day.
- **No real login yet** — the "Staff name" field is just a label typed in, not
  a password-protected account. Good enough to know who rang up a sale; not
  enough to stop someone from typing a different name. See improvement ideas
  below if you want real accounts.

## Ideas for later (not built yet, but straightforward additions)
- **Real staff logins** with different permissions (e.g. cashiers can't see profit).
- **CSV/PDF export** of reports for printing or sending to an accountant.
- **Barcode input** on the Sell page instead of picking from a dropdown.
- **Weighted-average buying price** if you often buy the same product at
  different prices — right now the app uses the most recent buying price for
  profit math, which is simpler but slightly less precise if prices swing a lot.
- **Custom date range** in Reports, beyond day/week/month.