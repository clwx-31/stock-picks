# Stock picks website

This repository owns https://clwx-31.github.io/stock-picks/. It is a separate
project from `../stock-weekly-tracker` (the earlier portfolio). "Update the
stocks" means this repository.

## Files

| File | Role |
|---|---|
| `portfolio.csv` | The original purchase on Oct 26, 2026: symbol, company, sector, shares, buy price. Source for the Picks tab and the Word write-up. |
| `history.csv` | The saved record of every week's prices, one row per stock per update. Append only. |
| `build.js` | Generates `index.html` (Picks), `weekly.html` (Weekly) and `Weekly-Archive.docx`, and checks the math. Never hand-edit those outputs. |
| `Weekly-Archive.docx` | Every week that has rolled off the Weekly tab, newest first. Created by the second update. |
| `tools/build_doc.js` | Rebuilds `Stock-Picks-Data-Centers.docx` from `portfolio.csv`. |

Run `npm install` once if `node_modules/` is missing; `build.js` needs it.

## How the Weekly tab works

- It shows exactly two prices per stock: last week's and this week's, each
  with its total (price × shares), then Difference = this week's total − last
  week's total, with a total row.
- Profit/loss started at $0.00 on Oct 26, 2026. Each update adds that week's
  total difference to last week's profit/loss, and the result is carried to the
  next week. Unspent budget (the rest of the $20,000–$25,000) is never counted.
- When a new week arrives, the week that drops off the page goes into
  `Weekly-Archive.docx`, which the Weekly tab links to. Older weeks are kept
  for reference only; they do not change the profit/loss.

## When the user says "update the stocks"

Update only the Weekly tab. The Picks tab, `portfolio.csv` and the Word
write-up record the original purchase and do not change.

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
   Never edit or delete earlier rows. If today already has rows, ask before
   replacing them.
4. Run `node build.js`. It moves this week's prices into next week's "Last
   week" columns automatically, computes each Difference, carries the
   profit/loss forward and archives the week that rolled off. It stops if a
   price is missing, dates are out of order, shares changed, or the
   profit/loss does not reconcile.
5. Recheck independently in integer cents: each row's totals and difference,
   the column totals, and last week's profit/loss + this week's difference =
   this week's profit/loss.
6. `git diff --stat` must show only `history.csv`, `weekly.html` and
   `Weekly-Archive.docx` (the archive appears from the second update on). Run
   `git diff --check`.
7. Commit as `Update stocks for <Mon D, YYYY>`, push to `origin main`, wait for
   the Pages build, and confirm the live `weekly.html` matches the local file.
8. Report the date and quote-time range, this week's total difference (gain or
   loss, and how much), the new profit/loss, and the biggest gain and loss.

## Trades

If the user buys or sells shares, do not fold it into a weekly update. Ask how
to record it first. Money spent on new shares must not show up as a gain, and
`build.js` refuses a share change between weeks until this is handled.

## Style

Black and white, minimal, with no colors added. Gains and losses are shown with
`+` and `−` signs. Keep both tabs working at phone width.
