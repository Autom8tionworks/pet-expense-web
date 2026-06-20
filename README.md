# PET Terra Systems — Expense Report Web App

A no-login web app: fill in a travel (job) or office expense report, optionally
scan receipts with Claude vision, and **download the completed Excel file**. It
fills your real fixed master templates server-side, so the downloaded workbook is
the official PET Terra form with all formulas intact.

* No email. No accounts. Anyone with the link enters their name and submits.
* Job and Office reports both supported.
* Receipt scanning is optional and powered by Claude vision.

---

## Project structure
```
index.html              the whole front-end (vanilla JS, mobile friendly)
api/
  generate.py           POST submission JSON -> downloads filled .xlsx
  scan.py               POST a receipt image -> Claude vision -> line item JSON
  _lib/                 shared engine (model, rows, filler, build) - validated
  _templates/           the two fixed masters (do not regenerate from originals)
requirements.txt        openpyxl + anthropic
vercel.json             bundles api/** with each function
```
The Excel-filling engine is the same one validated end-to-end against LibreOffice
recalculation (zero formula errors, correct totals, safe dynamic row insertion,
no XLOOKUP). In production no LibreOffice is needed — Excel recalculates the
formulas when the user opens the downloaded file.

---

## Deploy to Vercel from GitHub

1. **Put this folder in a GitHub repo.** From inside this folder:
   ```bash
   git init
   git add .
   git commit -m "PET Terra expense report web app"
   gh repo create pet-expense-report --private --source=. --push
   ```
   (Or create an empty repo on github.com and `git remote add origin <url>` then
   `git push -u origin main`.)

2. **Import into Vercel.**
   - Go to https://vercel.com/new and pick the GitHub repo.
   - Framework preset: **Other** (it's static + Python functions; no build step).
   - Leave Build/Output settings empty. Click **Deploy**.
   - Vercel auto-detects `api/*.py` and installs `requirements.txt`.

3. **Done.** You get a URL like `https://pet-expense-report.vercel.app`. Share it
   with the team. Every push to the repo redeploys automatically.

---

## Connect Claude (turn on receipt scanning)

Receipt scanning calls Claude vision through the Anthropic API. Without a key the
app still works for manual entry; the scan button just reports that scanning is
off.

1. **Get an API key** at https://console.anthropic.com → *API Keys* → create key.
2. **Add it to Vercel:** Project → **Settings → Environment Variables**:
   - `ANTHROPIC_API_KEY` = your key  (required)
   - `ANTHROPIC_MODEL`   = `claude-sonnet-4-6`  (optional; default already this).
     Use `claude-haiku-4-5-20251001` for cheaper/faster, or `claude-opus-4-8`
     for maximum accuracy.
3. **Redeploy** (Deployments → ⋯ → Redeploy) so the new env vars take effect.

That's the whole connection. `api/scan.py` reads the key from the environment,
sends the receipt image + an extraction prompt to Claude, and returns
`{vendor, date, amount, category, confidence}` which the page drops into a line
row (low-confidence rows are highlighted for review).

**Using a different AI provider instead of Claude:** edit `api/scan.py` — replace
the `anthropic` client block with your provider's vision call and map its output
to the same JSON shape. Swap the dependency in `requirements.txt` and add that
provider's API-key env var. Everything else (the prompt contract and the
front-end) stays the same.

> Cost & privacy note: receipt images are sent to the model provider only at scan
> time and are not stored by this app. Choosing where/whether to retain receipt
> images (financial PII) is still an open decision — see below.

---

## Run locally

```bash
npm i -g vercel        # one time
vercel dev             # serves the site + /api functions at localhost:3000
```
Set the env var locally with `vercel env pull` or a `.env` file
(`ANTHROPIC_API_KEY=...`) to test scanning.

---

## Notes, limits, open items

* **Mileage rate** ($0.68/mi) lives in the templates; making it a configurable
  setting is still a recommended follow-up (the rate changes periodically).
* **Row overflow** is handled automatically — extra mileage/expense/misc rows
  beyond the template's built-in rows are inserted safely.
* **Request size:** receipt photos are downscaled in the browser before upload to
  stay well under Vercel's request limit.
* **Still your decisions:** receipt-image retention policy, and confirming the
  "Total Section 3" template change with accounting before wide rollout.
