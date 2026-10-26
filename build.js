// Builds index.html (Picks), weekly.html (Weekly) and weekly-tracker.csv
// from portfolio.csv and history.csv. Run: node build.js
const fs = require('fs');
const path = require('path');

const DIR = __dirname;
const DOCX = 'Stock-Picks-Data-Centers.docx';
const BUDGET_MIN_C = 2000000;
const BUDGET_MAX_C = 2500000;

function readCsv(file) {
  const [head, ...lines] = fs.readFileSync(path.join(DIR, file), 'utf8').trim().split(/\r?\n/);
  const cols = head.split(',');
  return lines.map((line, i) => {
    const vals = line.split(',');
    if (vals.length !== cols.length) throw new Error(`${file} line ${i + 2}: expected ${cols.length} fields`);
    return Object.fromEntries(cols.map((c, j) => [c, vals[j].trim()]));
  });
}

function toCents(s, where) {
  if (!/^\d+(\.\d{1,2})?$/.test(s)) throw new Error(`${where}: bad price "${s}"`);
  const [whole, frac = ''] = s.split('.');
  return Number(whole) * 100 + Number(frac.padEnd(2, '0'));
}

function toInt(s, where) {
  if (!/^[1-9]\d*$/.test(s)) throw new Error(`${where}: bad share count "${s}"`);
  return Number(s);
}

const usd = (c) => '$' + (Math.abs(c) / 100).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const signed = (c) => (c > 0 ? '+' : c < 0 ? '−' : '') + usd(c);
const plain = (c) => (c < 0 ? '-' : '') + (Math.abs(c) / 100).toFixed(2);
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
const longDate = (iso) => new Date(iso + 'T12:00:00Z').toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric', timeZone: 'UTC' });
const shortDate = (iso) => new Date(iso + 'T12:00:00Z').toLocaleDateString('en-US', { month: 'short', day: 'numeric', timeZone: 'UTC' });

// "11:30 AM EDT" -> minutes after midnight, for showing each week's quote-time range.
function minutes(t, where) {
  const m = /^(\d{1,2}):(\d{2}) (AM|PM) E[SD]T$/.exec(t);
  if (!m) throw new Error(`${where}: bad quote_time "${t}"`);
  return (Number(m[1]) % 12 + (m[3] === 'PM' ? 12 : 0)) * 60 + Number(m[2]);
}
function timeRange(times) {
  const sorted = [...times].sort((a, b) => minutes(a) - minutes(b));
  const strip = (t) => t.replace(/ E[SD]T$/, '');
  const [lo, hi] = [strip(sorted[0]), strip(sorted.at(-1))];
  if (lo === hi) return `${lo} ET`;
  const [loT, loP] = lo.split(' ');
  const [hiT, hiP] = hi.split(' ');
  return loP === hiP ? `${loT}–${hiT} ${hiP} ET` : `${lo}–${hi} ET`;
}

// ---- holdings (the Picks tab: the original purchase, never changed by weekly updates) ----
const holdings = readCsv('portfolio.csv').map((r, i) => ({
  symbol: r.symbol, company: r.company, sector: r.sector,
  shares: toInt(r.shares, `portfolio.csv line ${i + 2}`),
  buyC: toCents(r.buy_price, `portfolio.csv line ${i + 2}`),
}));
holdings.forEach((h) => { h.costC = h.buyC * h.shares; });
const purchaseC = holdings.reduce((a, h) => a + h.costC, 0);
if (purchaseC < BUDGET_MIN_C || purchaseC > BUDGET_MAX_C) throw new Error(`purchase total ${usd(purchaseC)} is outside the budget range`);
const symbols = holdings.map((h) => h.symbol);
if (new Set(symbols).size !== symbols.length) throw new Error('duplicate symbol in portfolio.csv');

// ---- weekly snapshots ----
const byDate = new Map();
readCsv('history.csv').forEach((r, i) => {
  const where = `history.csv line ${i + 2}`;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(r.date)) throw new Error(`${where}: bad date "${r.date}"`);
  if (!symbols.includes(r.symbol)) throw new Error(`${where}: ${r.symbol} is not in portfolio.csv`);
  minutes(r.quote_time, where);
  if (!byDate.has(r.date)) byDate.set(r.date, new Map());
  const day = byDate.get(r.date);
  if (day.has(r.symbol)) throw new Error(`${where}: ${r.symbol} appears twice on ${r.date}`);
  day.set(r.symbol, { shares: toInt(r.shares, where), priceC: toCents(r.price, where), time: r.quote_time });
});
const dates = [...byDate.keys()];
if (dates.join() !== [...dates].sort().join()) throw new Error('history.csv dates must be in order, oldest first');

const weeks = dates.map((date, i) => {
  const day = byDate.get(date);
  const missing = symbols.filter((s) => !day.has(s));
  if (missing.length) throw new Error(`${date} is missing prices for ${missing.join(', ')}`);
  const prevDate = i > 0 ? dates[i - 1] : null;
  const prev = prevDate ? byDate.get(prevDate) : null;
  const rows = holdings.map((h) => {
    const cur = day.get(h.symbol);
    const last = prev ? prev.get(h.symbol) : null;
    if (last && last.shares !== cur.shares) {
      throw new Error(`${h.symbol} shares changed from ${last.shares} to ${cur.shares} on ${date}; record trades as AGENTS.md describes`);
    }
    const thisC = cur.priceC * cur.shares;
    const lastC = last ? last.priceC * last.shares : null;
    return { h, shares: cur.shares, lastPriceC: last ? last.priceC : null, lastC, priceC: cur.priceC, thisC, diffC: last ? thisC - lastC : null };
  });
  const thisC = rows.reduce((a, r) => a + r.thisC, 0);
  const lastC = prev ? rows.reduce((a, r) => a + r.lastC, 0) : null;
  const diffC = prev ? rows.reduce((a, r) => a + r.diffC, 0) : null;
  return { date, prevDate, rows, thisC, lastC, diffC, plC: thisC - purchaseC, times: timeRange([...day.values()].map((v) => v.time)) };
});

// Week 1 is the purchase itself, so profit/loss starts at exactly $0.00.
const first = byDate.get(dates[0]);
holdings.forEach((h) => {
  const r = first.get(h.symbol);
  if (r.priceC !== h.buyC || r.shares !== h.shares) throw new Error(`${dates[0]} ${h.symbol} must match portfolio.csv`);
});
weeks.forEach((w, i) => {
  if (i === 0 && w.plC !== 0) throw new Error('first week profit/loss must be $0.00');
  if (i > 0 && w.plC !== weeks[i - 1].plC + w.diffC) throw new Error(`${w.date}: profit/loss does not reconcile with last week plus this week's difference`);
});

// ---- shared page parts ----
const head = (title, description) => `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${title}</title>
  <meta name="description" content="${description}">
  <style>
    :root { --fg: #000; --bg: #fff; --muted: #6e6e6e; --rule: #e6e6e6; }
    @media (prefers-color-scheme: dark) {
      :root { --fg: #fff; --bg: #000; --muted: #9a9a9a; --rule: #262626; }
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    html { -webkit-text-size-adjust: 100%; }
    body {
      background: var(--bg);
      color: var(--fg);
      font: 16px/1.5 -apple-system, BlinkMacSystemFont, "Helvetica Neue", Helvetica, Arial, sans-serif;
      -webkit-font-smoothing: antialiased;
    }
    main { max-width: 760px; margin: 0 auto; padding: 64px 24px 96px; }
    main.wide { max-width: 1040px; }
    .tabs { display: flex; gap: 28px; margin-bottom: 48px; font-size: 15px; }
    .tabs a { color: var(--muted); text-decoration: none; padding-bottom: 4px; border-bottom: 1px solid transparent; }
    .tabs a:hover { color: var(--fg); }
    .tabs a[aria-current="page"] { color: var(--fg); border-bottom-color: var(--fg); }
    h1 { font-size: 32px; font-weight: 600; letter-spacing: -0.025em; line-height: 1.15; }
    .meta { margin-top: 10px; color: var(--muted); font-size: 15px; }
    .table-wrap { margin-top: 56px; overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; font-variant-numeric: tabular-nums; }
    th {
      padding: 0 0 12px;
      border-bottom: 1px solid var(--fg);
      color: var(--muted);
      font-size: 12px;
      font-weight: 500;
      letter-spacing: 0.06em;
      text-align: left;
      text-transform: uppercase;
      vertical-align: bottom;
    }
    td { padding: 15px 0; border-bottom: 1px solid var(--rule); vertical-align: top; }
    th + th, td + td { padding-left: 20px; }
    .num { text-align: right; white-space: nowrap; }
    .sym { font-weight: 500; letter-spacing: 0.02em; }
    tfoot td { border-top: 1px solid var(--fg); border-bottom: 0; font-weight: 600; }
    .note { margin-top: 16px; color: var(--muted); font-size: 14px; }
    .download {
      display: inline-block;
      margin-top: 56px;
      padding: 13px 22px;
      border: 1px solid var(--fg);
      color: var(--fg);
      font-size: 15px;
      text-decoration: none;
      transition: background 0.15s, color 0.15s;
    }
    .download:hover, .download:focus-visible { background: var(--fg); color: var(--bg); outline: none; }
    .download span { color: var(--muted); }
    .download:hover span, .download:focus-visible span { color: inherit; }
    .picker { display: flex; align-items: center; gap: 12px; margin-top: 40px; font-size: 15px; }
    .picker label { color: var(--muted); }
    .picker select {
      appearance: none;
      -webkit-appearance: none;
      background: var(--bg);
      color: var(--fg);
      border: 1px solid var(--fg);
      border-radius: 0;
      font: inherit;
      padding: 8px 36px 8px 12px;
      background-image: linear-gradient(45deg, transparent 50%, var(--fg) 50%), linear-gradient(135deg, var(--fg) 50%, transparent 50%);
      background-position: calc(100% - 18px) 50%, calc(100% - 13px) 50%;
      background-size: 5px 5px;
      background-repeat: no-repeat;
    }
    .week .table-wrap { margin-top: 32px; }
    .week table { min-width: 860px; }
    th.group { text-align: center; padding-bottom: 8px; border-bottom: 1px solid var(--rule); }
    thead tr + tr th { padding-top: 8px; }
    .empty { color: var(--muted); }
    .pl { margin-top: 48px; padding-top: 24px; border-top: 1px solid var(--fg); }
    .pl-label { color: var(--muted); font-size: 12px; font-weight: 500; letter-spacing: 0.06em; text-transform: uppercase; }
    .pl-value { margin-top: 6px; font-size: 40px; font-weight: 600; letter-spacing: -0.03em; font-variant-numeric: tabular-nums; }
    @media (max-width: 600px) {
      main { padding: 32px 16px 64px; }
      .tabs { margin-bottom: 32px; }
      h1 { font-size: 26px; }
      .table-wrap { margin-top: 40px; }
      table { font-size: 14px; }
      th { font-size: 11px; letter-spacing: 0.04em; }
      th + th, td + td { padding-left: 10px; }
      td { padding: 12px 0; }
      .week table { min-width: 680px; }
      .pl-value { font-size: 32px; }
      .download { display: block; text-align: center; margin-top: 40px; }
    }
  </style>
</head>`;

const tabs = (current) => `    <nav class="tabs" aria-label="Pages">
      <a href="./"${current === 'picks' ? ' aria-current="page"' : ''}>Picks</a>
      <a href="weekly.html"${current === 'weekly' ? ' aria-current="page"' : ''}>Weekly</a>
    </nav>`;

// ---- Picks tab ----
const sectorCount = new Set(holdings.map((h) => h.sector)).size;
const kb = Math.round(fs.statSync(path.join(DIR, DOCX)).size / 1024);
const picksRows = holdings.map((h) => `        <tr>
          <td>${esc(h.company)}</td>
          <td class="sym">${h.symbol}</td>
          <td class="num">${usd(h.buyC)}</td>
          <td class="num">${h.shares}</td>
          <td class="num">${usd(h.costC)}</td>
        </tr>`).join('\n');

const picks = `${head('Stock Picks', `${holdings.length} stock picks for a business class portfolio, with price, shares and total cost.`)}
<body>
  <main>
${tabs('picks')}
    <h1>Stock Picks</h1>
    <p class="meta">${holdings.length} companies &middot; ${sectorCount} sectors &middot; Prices as of ${longDate(dates[0])}, ${weeks[0].times.replace(/^(\S+)–\S+ /, '$1 ')}</p>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th scope="col">Company</th>
            <th scope="col">Symbol</th>
            <th scope="col" class="num">Price per share</th>
            <th scope="col" class="num">Shares</th>
            <th scope="col" class="num">Total cost</th>
          </tr>
        </thead>
        <tbody>
${picksRows}
        </tbody>
        <tfoot>
          <tr>
            <td>Total</td>
            <td></td>
            <td></td>
            <td></td>
            <td class="num">${usd(purchaseC)}</td>
          </tr>
        </tfoot>
      </table>
    </div>
    <p class="note">Budget $20,000 to $25,000</p>

    <a class="download" href="${DOCX}" download>Download the full write-up <span>(Word, ${kb} KB)</span></a>
  </main>
</body>
</html>
`;

// ---- Weekly tab ----
const dash = '<span class="empty">—</span>';
const cellC = (c, fmt = usd) => (c === null ? dash : fmt(c));
const newestFirst = [...weeks].reverse();

const weekSection = (w, i) => {
  const lastLabel = w.prevDate ? `Last week &middot; ${shortDate(w.prevDate)}` : 'Last week';
  const body = w.rows.map((r) => `            <tr>
              <td>${esc(r.h.company)}</td>
              <td class="sym">${r.h.symbol}</td>
              <td class="num">${r.shares}</td>
              <td class="num">${cellC(r.lastPriceC)}</td>
              <td class="num">${cellC(r.lastC)}</td>
              <td class="num">${usd(r.priceC)}</td>
              <td class="num">${usd(r.thisC)}</td>
              <td class="num">${cellC(r.diffC, signed)}</td>
            </tr>`).join('\n');
  const compared = w.prevDate ? `compared with ${longDate(w.prevDate)}` : 'starting week, so there is no last week yet';
  return `    <section class="week" data-week="${w.date}"${i === 0 ? '' : ' hidden'}>
      <p class="meta">Prices as of ${longDate(w.date)}, ${w.times} &middot; ${compared}</p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th scope="col" rowspan="2">Company</th>
              <th scope="col" rowspan="2">Symbol</th>
              <th scope="col" rowspan="2" class="num">Shares</th>
              <th scope="colgroup" colspan="2" class="group">${lastLabel}</th>
              <th scope="colgroup" colspan="2" class="group">This week &middot; ${shortDate(w.date)}</th>
              <th scope="col" rowspan="2" class="num">Difference</th>
            </tr>
            <tr>
              <th scope="col" class="num">Price</th>
              <th scope="col" class="num">Total</th>
              <th scope="col" class="num">Price</th>
              <th scope="col" class="num">Total</th>
            </tr>
          </thead>
          <tbody>
${body}
          </tbody>
          <tfoot>
            <tr>
              <td>Total</td>
              <td></td>
              <td></td>
              <td></td>
              <td class="num">${cellC(w.lastC)}</td>
              <td></td>
              <td class="num">${usd(w.thisC)}</td>
              <td class="num">${cellC(w.diffC, signed)}</td>
            </tr>
          </tfoot>
        </table>
      </div>
      <div class="pl">
        <p class="pl-label">Profit / loss since purchase</p>
        <p class="pl-value">${signed(w.plC)}</p>
        <p class="note">${usd(w.thisC)} now vs. ${usd(purchaseC)} paid on ${longDate(dates[0])}</p>
      </div>
    </section>`;
};

const options = newestFirst.map((w, i) => `        <option value="${w.date}"${i === 0 ? ' selected' : ''}>${longDate(w.date)}${w.prevDate ? '' : ' (start)'}</option>`).join('\n');

const weekly = `${head('Weekly Tracker', 'Week-by-week prices, totals and profit or loss for the stock picks portfolio.')}
<body>
  <main class="wide">
${tabs('weekly')}
    <h1>Weekly Tracker</h1>
    <p class="meta">Live prices recorded at the moment of each weekly update</p>

    <div class="picker">
      <label for="week">Week</label>
      <select id="week">
${options}
      </select>
    </div>

${newestFirst.map(weekSection).join('\n\n')}

    <a class="download" href="weekly-tracker.csv" download>Download every week <span>(CSV)</span></a>
  </main>
  <script>
    const select = document.getElementById('week');
    const show = () => document.querySelectorAll('.week').forEach((s) => { s.hidden = s.dataset.week !== select.value; });
    select.addEventListener('change', show);
    show();
  </script>
</body>
</html>
`;

// ---- spreadsheet export: every week, every row, plus a total row per week ----
const csvCell = (v) => (/[",]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);
const csv = [['week', 'compared_with', 'company', 'symbol', 'shares', 'last_week_price', 'last_week_total', 'this_week_price', 'this_week_total', 'difference', 'profit_loss_since_purchase']];
weeks.forEach((w) => {
  w.rows.forEach((r) => csv.push([w.date, w.prevDate || '', r.h.company, r.h.symbol, String(r.shares),
    r.lastPriceC === null ? '' : plain(r.lastPriceC), r.lastC === null ? '' : plain(r.lastC),
    plain(r.priceC), plain(r.thisC), r.diffC === null ? '' : plain(r.diffC), '']));
  csv.push([w.date, w.prevDate || '', 'Total', '', '', '', w.lastC === null ? '' : plain(w.lastC), '', plain(w.thisC),
    w.diffC === null ? '' : plain(w.diffC), plain(w.plC)]);
});

fs.writeFileSync(path.join(DIR, 'index.html'), picks);
fs.writeFileSync(path.join(DIR, 'weekly.html'), weekly);
fs.writeFileSync(path.join(DIR, 'weekly-tracker.csv'), csv.map((r) => r.map(csvCell).join(',')).join('\n') + '\n');

const latest = weeks.at(-1);
console.log(`purchase ${usd(purchaseC)} on ${dates[0]}`);
console.log(`latest week ${latest.date} (${latest.times})${latest.prevDate ? ` vs ${latest.prevDate}` : ''}`);
console.log(`  this week ${usd(latest.thisC)}  difference ${latest.diffC === null ? 'n/a' : signed(latest.diffC)}  profit/loss ${signed(latest.plC)}`);
