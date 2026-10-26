// Builds index.html (Picks), weekly.html (Weekly) and Weekly-Archive.docx
// from portfolio.csv and history.csv. Run: npm install (once), then node build.js
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, PageOrientation, TableLayoutType,
} = require('docx');

const DIR = __dirname;
const DOCX = 'Stock-Picks-Data-Centers.docx';
const ARCHIVE = 'Weekly-Archive.docx';
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
  return { date, prevDate, rows, thisC, lastC, diffC, times: timeRange([...day.values()].map((v) => v.time)) };
});

// The first week is the purchase itself.
const first = byDate.get(dates[0]);
holdings.forEach((h) => {
  const r = first.get(h.symbol);
  if (r.priceC !== h.buyC || r.shares !== h.shares) throw new Error(`${dates[0]} ${h.symbol} must match portfolio.csv`);
});

// Profit/loss starts at $0.00 and carries forward: each week adds its difference
// to last week's saved amount. Unspent budget is never counted.
weeks.forEach((w, i) => { w.plC = i === 0 ? 0 : weeks[i - 1].plC + w.diffC; });
// Cross-check: with shares unchanged, the carried amount equals stock value minus purchase cost.
weeks.forEach((w) => { if (w.plC !== w.thisC - purchaseC) throw new Error(`${w.date}: profit/loss does not reconcile`); });

const plNote = (w, i) => (i === 0
  ? 'Starting point. Each update adds that week’s difference.'
  : `Last week’s ${signed(weeks[i - 1].plC)} plus this week’s difference of ${signed(w.diffC)}`);

// The website shows only the latest week. Earlier weeks (after the starting week,
// whose prices appear as week 2's "last week") roll off into the Word archive.
const latestIndex = weeks.length - 1;
const latest = weeks[latestIndex];
const archived = weeks.slice(1, -1).map((w, k) => ({ w, i: k + 1 })).reverse();

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
    .week table { min-width: 860px; }
    th.group { text-align: center; padding-bottom: 8px; border-bottom: 1px solid var(--rule); }
    thead tr + tr th { padding-top: 8px; }
    .empty { color: var(--muted); }
    .pl { margin-top: 48px; padding-top: 24px; border-top: 1px solid var(--fg); }
    .pl-label { color: var(--muted); font-size: 12px; font-weight: 500; letter-spacing: 0.06em; text-transform: uppercase; }
    .pl-value { margin-top: 6px; font-size: 40px; font-weight: 600; letter-spacing: -0.03em; font-variant-numeric: tabular-nums; }
    .later { margin-top: 56px; }
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
const kb = (file) => Math.round(fs.statSync(path.join(DIR, file)).size / 1024);
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

    <a class="download" href="${DOCX}" download>Download the full write-up <span>(Word, ${kb(DOCX)} KB)</span></a>
  </main>
</body>
</html>
`;

// ---- Weekly tab: the latest week only ----
const dash = '<span class="empty">—</span>';
const cellC = (c, fmt = usd) => (c === null ? dash : fmt(c));
const lastLabel = latest.prevDate ? `Last week &middot; ${shortDate(latest.prevDate)}` : 'Last week';
const weeklyRows = latest.rows.map((r) => `            <tr>
              <td>${esc(r.h.company)}</td>
              <td class="sym">${r.h.symbol}</td>
              <td class="num">${r.shares}</td>
              <td class="num">${cellC(r.lastPriceC)}</td>
              <td class="num">${cellC(r.lastC)}</td>
              <td class="num">${usd(r.priceC)}</td>
              <td class="num">${usd(r.thisC)}</td>
              <td class="num">${cellC(r.diffC, signed)}</td>
            </tr>`).join('\n');
const compared = latest.prevDate ? `compared with ${longDate(latest.prevDate)}` : 'starting week, so there is no last week yet';
const archiveBlock = archived.length
  ? `    <a class="download" href="${ARCHIVE}" download>Download past weeks <span>(Word)</span></a>`
  : '    <p class="note later">Past weeks are saved to a Word document once they roll off this page.</p>';

const weekly = `${head('Weekly Tracker', 'This week and last week’s prices, totals and profit or loss for the stock picks portfolio.')}
<body>
  <main class="wide">
${tabs('weekly')}
    <h1>Weekly Tracker</h1>
    <p class="meta">Prices as of ${longDate(latest.date)}, ${latest.times} &middot; ${compared}</p>

    <section class="week">
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th scope="col" rowspan="2">Company</th>
              <th scope="col" rowspan="2">Symbol</th>
              <th scope="col" rowspan="2" class="num">Shares</th>
              <th scope="colgroup" colspan="2" class="group">${lastLabel}</th>
              <th scope="colgroup" colspan="2" class="group">This week &middot; ${shortDate(latest.date)}</th>
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
${weeklyRows}
          </tbody>
          <tfoot>
            <tr>
              <td>Total</td>
              <td></td>
              <td></td>
              <td></td>
              <td class="num">${cellC(latest.lastC)}</td>
              <td></td>
              <td class="num">${usd(latest.thisC)}</td>
              <td class="num">${cellC(latest.diffC, signed)}</td>
            </tr>
          </tfoot>
        </table>
      </div>
      <div class="pl">
        <p class="pl-label">Profit / loss</p>
        <p class="pl-value">${signed(latest.plC)}</p>
        <p class="note">${plNote(latest, latestIndex)}</p>
      </div>
    </section>

${archiveBlock}
  </main>
</body>
</html>
`;

// ---- Word archive of weeks that have rolled off the website ----
const FONT = 'Calibri';
const INK = '000000';
const GRAY = '6E6E6E';
const LINE = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const DARK = { style: BorderStyle.SINGLE, size: 8, color: INK };
const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const COLS = [3000, 1100, 1000, 1450, 1650, 1450, 1650, 1660]; // 12,960 DXA = 9" landscape content width

const docRun = (text, o = {}) => new TextRun({ text, font: FONT, color: INK, ...o });
function docCell(text, i, o = {}) {
  const right = i >= 2;
  return new TableCell({
    width: { size: COLS[i], type: WidthType.DXA },
    borders: { top: o.top || NONE, bottom: o.bottom || LINE, left: NONE, right: NONE },
    margins: { top: 50, bottom: 50, left: i === 0 ? 0 : 80, right: 0 },
    children: [new Paragraph({
      alignment: right ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [docRun(text, { bold: o.bold, color: o.color || INK, size: o.size || 19 })],
    })],
  });
}

function archiveWeek({ w, i }, first) {
  const header = ['Company', 'Symbol', 'Shares', `${shortDate(w.prevDate)} price`, `${shortDate(w.prevDate)} total`,
    `${shortDate(w.date)} price`, `${shortDate(w.date)} total`, 'Difference'];
  const rows = [
    new TableRow({ tableHeader: true, children: header.map((t, c) => docCell(t, c, { bold: true, color: GRAY, size: 17, bottom: DARK })) }),
    ...w.rows.map((r) => new TableRow({ children: [r.h.company, r.h.symbol, String(r.shares), usd(r.lastPriceC), usd(r.lastC),
      usd(r.priceC), usd(r.thisC), signed(r.diffC)].map((t, c) => docCell(t, c)) })),
    new TableRow({ children: ['Total', '', '', '', usd(w.lastC), '', usd(w.thisC), signed(w.diffC)]
      .map((t, c) => docCell(t, c, { bold: true, top: DARK, bottom: NONE })) }),
  ];
  return [
    new Paragraph({ pageBreakBefore: !first, spacing: { before: first ? 0 : 0, after: 60 },
      children: [docRun(`Week of ${longDate(w.date)}`, { bold: true, size: 30 })] }),
    new Paragraph({ spacing: { after: 200 },
      children: [docRun(`Prices as of ${longDate(w.date)}, ${w.times} · compared with ${longDate(w.prevDate)}`, { color: GRAY, size: 19 })] }),
    new Table({
      width: { size: 12960, type: WidthType.DXA }, columnWidths: COLS, layout: TableLayoutType.FIXED, rows,
      borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE },
    }),
    new Paragraph({ spacing: { before: 280, after: 40 },
      children: [docRun('Profit / loss: ', { bold: true, size: 24 }), docRun(signed(w.plC), { bold: true, size: 24 })] }),
    new Paragraph({ children: [docRun(plNote(w, i), { color: GRAY, size: 19 })] }),
  ];
}

async function writeArchive() {
  const file = path.join(DIR, ARCHIVE);
  if (!archived.length) {
    if (fs.existsSync(file)) throw new Error(`${ARCHIVE} exists but no weeks have rolled off yet`);
    return 'no past weeks yet';
  }
  const doc = new Document({
    creator: 'Stock project',
    title: 'Weekly Tracker Archive',
    styles: { default: { document: { run: { font: FONT, size: 20, color: INK } } } },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840, orientation: PageOrientation.LANDSCAPE },
        margin: { top: 1080, bottom: 1080, left: 1440, right: 1440 } } },
      children: [
        new Paragraph({ spacing: { after: 60 }, children: [docRun('Weekly Tracker Archive', { bold: true, size: 40 })] }),
        new Paragraph({ spacing: { after: 360 }, children: [docRun(
          'Weeks that have rolled off the website, newest first. Profit/loss carries forward: each week adds its difference to the week before.',
          { color: GRAY, size: 19 })] }),
        ...archived.flatMap((a, k) => archiveWeek(a, k === 0)),
      ],
    }],
  });
  const buf = await Packer.toBuffer(doc);
  // Skip rewriting when only the embedded timestamps would change.
  if (fs.existsSync(file)) {
    const xml = async (b) => (await JSZip.loadAsync(b)).file('word/document.xml').async('string');
    if (await xml(buf) === await xml(fs.readFileSync(file))) return 'unchanged';
  }
  fs.writeFileSync(file, buf);
  return `${archived.length} past week${archived.length === 1 ? '' : 's'}`;
}

(async () => {
  fs.writeFileSync(path.join(DIR, 'index.html'), picks);
  fs.writeFileSync(path.join(DIR, 'weekly.html'), weekly);
  const archiveStatus = await writeArchive();
  console.log(`purchase ${usd(purchaseC)} on ${dates[0]}`);
  console.log(`weekly tab ${latest.date} (${latest.times})${latest.prevDate ? ` vs ${latest.prevDate}` : ''}`);
  console.log(`  this week ${usd(latest.thisC)}  difference ${latest.diffC === null ? 'n/a' : signed(latest.diffC)}  profit/loss ${signed(latest.plC)}`);
  console.log(`archive: ${archiveStatus}`);
})().catch((e) => { console.error(e.message); process.exit(1); });
