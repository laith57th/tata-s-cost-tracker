# Tata's Batters — Cost Tracker Web App

A web app that reads and writes your Google Sheet directly. Add recipes, ingredients,
products, and update prices from your phone or laptop without opening the spreadsheet.

Runs entirely inside your own Google account. No hosting, no API keys, no monthly cost.

---

## Setup (about 5 minutes, one time)

**1. Open your sheet**

Open `Tatas_Batters_Menu_Pricing` (the cost tracker) in Google Sheets.

**2. Open the script editor**

Menu: **Extensions ▸ Apps Script**

A new tab opens with a file called `Code.gs` containing a stub `myFunction()`.

**3. Paste the backend**

Select everything in `Code.gs`, delete it, and paste in the full contents of the
`Code.gs` file from this folder.

**4. Add the interface file**

Click the **+** next to "Files" ▸ **HTML**. Name it exactly:

```
Index
```

(Apps Script adds the `.html` itself — don't type the extension.)

Delete the placeholder contents and paste in the full contents of `Index.html`.

**5. Save**

Click the save icon (or Ctrl/Cmd + S).

**6. Deploy**

- Click **Deploy ▸ New deployment**
- Click the gear icon next to "Select type" and choose **Web app**
- Fill in:
  - **Description:** anything, e.g. "Cost tracker v1"
  - **Execute as:** *Me*
  - **Who has access:** *Only myself*
- Click **Deploy**

**7. Authorize**

Google will ask for permission the first time.

- Click **Authorize access**, choose your account
- You'll likely see **"Google hasn't verified this app"** — this is expected for a
  script you wrote yourself. Click **Advanced**, then **Go to (project name) (unsafe)**,
  then **Allow**.
- This warning appears because the app isn't publicly published to Google's
  marketplace. It's your own code running on your own sheet.

**8. Copy your URL**

You'll get a link ending in `/exec`. That's your app.

Bookmark it. On a phone: open it in Chrome/Safari, then "Add to Home Screen" and it
behaves like an app.

---

## What each tab does

**Add Recipe Line** — Pick a product once, then add ingredients one after another.
The product stays selected, so entering a 7-ingredient recipe means picking the
product once instead of seven times. Live cost and running total shown below, with a
remove button per line.

**Update Price** — Pick an ingredient, see its current price and cost-per-unit, type
the new number. Every recipe using it recalculates automatically. Stamps today's date
into "Last Updated" for you.

**New Ingredient** — Name, Arabic name, type, purchase size/price/qty, unit. The
cost-per-unit formula is written in automatically. Blocks duplicates.

**New Product** — Adds the product to the summary sheet with all its formulas, and
starts a recipe block for it so you can add ingredients immediately.

**View Costs** — Read-only. Product costs and margins, plus a searchable/filterable
ingredient list.

---

## Notes

- **Everything writes to your real sheet.** The spreadsheet stays the source of truth;
  this is just a friendlier way to edit it. You can still open the sheet directly
  anytime.
- **Recipe lines reuse blank rows** inside a product's block first. If a block is full,
  the app inserts a new row and copies the formulas down, so the structure stays intact.
- **Don't delete column J** ("Product (resolved)") on Recipe Costing — both the sheet
  and this app depend on it.
- **After changing the code**, redeploy with **Deploy ▸ Manage deployments ▸ (pencil
  icon) ▸ Version: New version ▸ Deploy**. Just saving isn't enough.
- **"Only myself" access** means you must be signed into the Google account that owns
  the sheet. To let someone else use it, change access to "Anyone with the link" —
  but note that "Execute as: Me" means they'd be editing *your* sheet.
