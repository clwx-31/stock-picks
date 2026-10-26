# Stock Picks

Published website: https://clwx-31.github.io/stock-picks/

A minimal, black-and-white site for the business class portfolio.

- **Picks** tab: the original purchase on Oct 26, 2026. It lists the company,
  symbol, price per share, shares and total cost for all 13 stocks, and links
  the full Word write-up (`Stock-Picks-Data-Centers.docx`). Total invested is
  $21,043.64, inside the $20,000–$25,000 budget.
- **Weekly** tab: last week's and this week's price and total for each stock,
  the difference, and the profit or loss. Profit/loss started at $0.00 and each
  week adds that week's difference. Weeks that roll off the page are saved in
  `Weekly-Archive.docx`, linked from the tab.

Prices are live Stock Analysis quotes taken at the moment of each update.

## Updating

Say "update the stocks". [AGENTS.md](AGENTS.md) holds the full procedure: add
this week's prices to `history.csv`, run `node build.js`, then commit and push.
Older weeks are never edited.

| File | Purpose |
|---|---|
| `portfolio.csv` | Original holdings and purchase prices |
| `history.csv` | Every week's prices (append only) |
| `build.js` | Generates both pages and the weekly archive, and checks the math (`npm install` first) |
| `tools/build_doc.js` | Rebuilds the Word write-up |
