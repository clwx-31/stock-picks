# Stock picks website

This repository owns https://clwx-31.github.io/stock-picks/. It is a separate
project from `../stock-weekly-tracker` (the earlier portfolio). "Update the
stocks" means this repository.

## Files

| File | Role |
|---|---|
| `portfolio.csv` | The original purchase on Oct 26, 2026: symbol, company, sector, shares, buy price. Source for the Picks tab and the Word doc. |
| `history.csv` | The saved record of every week's prices, one row per stock per update. Append only. |
| `build.js` | Generates `index.html` (Picks), `weekly.html` (Weekly) and `weekly-tracker.csv`, and checks the math. Never hand-edit those three outputs. |
| `tools/build_doc.js` | Rebuilds `Stock-Picks-Data-Centers.docx` from `portfolio.csv` (`npm install` first). |

## When the user says "update the stocks"

Update only the Weekly tab. The Picks tab, `portfolio.csv` and the Word doc
record the original purchase and do not change.

1. Check the actual current date and time in ET. The latest date in
   `history.csv` is last week.
2. Fetch live quotes for all 13 symbols in `portfolio.csv` the moment the user
   asks, in parallel so the times are close together:
   `https://api.stockanalysis.com/api/quotes/s/<symbol in lowercase>`.
   Use `p` as the price, rounded half up to the cent (173.525 → 173.53), and
   the time from `u` (for example `11:33 AM EST`). Do not use closing prices or
   the cached stockanalysis.com HTML pages. If the market is closed, the quote
   is the last trade; say so in the report.
3. Append 13 rows to `history.csv` as `date,symbol,shares,price,quote_time`,
   dated today (`YYYY-MM-DD`), with shares carried forward from last week.
   Never edit or delete earlier rows: they are how the user looks back at old
   weeks. If today already has rows, ask before replacing them.
4. Run `node build.js`. It pairs each week with the week before it, so this
   week's prices become next week's "Last week" columns automatically. For
   each stock, Difference = this week's total − last week's total, and the
   total is price × shares. Profit/loss since purchase = this week's total −
   $21,043.64. The build stops if a price is missing, dates are out of order,
   shares changed, or profit/loss does not reconcile.
5. Recheck independently in integer cents: each row's totals and difference,
   the column totals, and last week's profit/loss + this week's difference =
   this week's profit/loss.
6. `git diff --stat` must show only `history.csv`, `weekly.html` and
   `weekly-tracker.csv`. Run `git diff --check`.
7. Commit as `Update stocks for <Mon D, YYYY>`, push to `origin main`, wait for
   the Pages build, and confirm the live `weekly.html` matches the local file.
8. Report the date and quote-time range, this week's total, this week's
   difference, profit/loss since purchase, and the biggest gain and loss.

## Trades

If the user buys or sells shares, do not fold it into a weekly update. Ask how
to record it first. Money spent on new shares must not show up as a gain, and
`build.js` refuses a share change between weeks until this is handled.

## Style

Black and white, minimal, with no colors added. Gains and losses are shown with
`+` and `−` signs. Keep both tabs working at phone width.
