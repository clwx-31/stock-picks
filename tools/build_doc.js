const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, AlignmentType, HeadingLevel, LevelFormat, BorderStyle, Footer,
  PageNumber, ExternalHyperlink, TableLayoutType,
} = require('docx');

// Builds the Word write-up. Shares and prices come from portfolio.csv.
// Run from stocks/: npm install && node tools/build_doc.js
const path = require('path');
const ROOT = path.join(__dirname, '..');
const OUT = process.argv[2] || path.join(ROOT, 'Stock-Picks-Data-Centers.docx');
const BUDGET_MIN = 20000;
const BUDGET_MAX = 25000;

// Prices: Stock Analysis live quotes, Oct 26 2026, 11:30-11:34 AM EDT, rounded to the cent.
const H = [
  {
    t: 'EME', n: 'EMCOR Group', s: 'Industrials', build: true,
    cap: '$29.5B', fpe: '19.3', div: '0.24%', beta: '1.12', earn: 'Oct 29', hi: 951.96, lo: 564.92,
    what: 'One of the largest electrical and mechanical construction contractors in the U.S. EMCOR installs the wiring, power distribution, cooling, plumbing and fire protection inside large buildings: data centers, hospitals, factories and labs.',
    why: 'Electrical and cooling systems are the most expensive part of building a data center, and that is EMCOR’s core work. Its Q2 2026 revenue rose 19.8% to $5.15 billion, earnings per share rose 34.8% to $9.06, and it raised its full-year guidance. Its backlog of signed work hit a record $17.14 billion, up 43.9% in a year, with the biggest gains in the segment that includes data centers. At about 19 times next year’s earnings, it is the best-value builder in this portfolio, which is why it is also the largest position.',
    risk: 'Many contracts are fixed-price, so cost overruns come out of EMCOR’s profit. The stock is about 30% below its 52-week high because of the October selloff in data-center stocks.',
  },
  {
    t: 'PWR', n: 'Quanta Services', s: 'Industrials', build: true,
    cap: '$85.6B', fpe: '30.6', div: '0.08%', beta: '1.27', earn: 'Oct 29', hi: 788.75, lo: 408.36,
    what: 'The largest U.S. contractor for electric power infrastructure. Quanta builds and maintains transmission lines, substations, and the grid connections that bring power to a site.',
    why: 'Analysts call electricity the biggest bottleneck for new data centers. A finished building is useless until it is connected to the grid, and Quanta does that work. It also benefits from grid upgrades that have nothing to do with AI, such as replacing aging lines and hardening against storms, which gives it demand even if AI spending slows.',
    risk: 'This is the most expensive stock in the portfolio: 65 times last year’s earnings and about 31 times next year’s. Its market value fell about $25 billion after October 6. It reports earnings on October 29.',
  },
  {
    t: 'STRL', n: 'Sterling Infrastructure', s: 'Industrials', build: true,
    cap: '$11.3B', fpe: '16.5', div: 'None', beta: '1.86', earn: 'Oct 28', hi: 1005.68, lo: 281.58,
    what: 'A site-development contractor. Sterling clears and grades land, pours foundations, and lays underground utilities for data centers, warehouses and factories. It also builds roads and bridges.',
    why: 'This is the most direct “physical construction” stock in the portfolio. Sterling is usually the first company on a data-center site. After the selloff it trades at about 16.5 times next year’s earnings, which is cheap for a company that has been growing this fast.',
    risk: 'It is the most volatile pick here, with a beta of 1.86, so it moves almost twice as much as the market. It fell 9.4% on October 22 alone and is about 63% below its 52-week high. Its position is kept small on purpose. It reports earnings on October 28.',
  },
  {
    t: 'VMC', n: 'Vulcan Materials', s: 'Materials', build: true,
    cap: '$34.0B', fpe: '28.4', div: '0.79%', beta: '1.05', earn: 'Oct 29', hi: 331.09, lo: 237.03,
    what: 'The largest U.S. producer of construction aggregates (crushed stone, sand and gravel), plus asphalt and ready-mix concrete.',
    why: 'Every data center sits on tons of concrete, and concrete is mostly aggregate. Stone is heavy and expensive to truck, so a quarry close to a job site faces little competition and has strong pricing power. Much of Vulcan’s demand comes from highways and public infrastructure, which run on a different cycle than AI spending. That makes it a theme stock that also adds diversification.',
    risk: 'High interest rates slow construction, and bad weather delays shipments. It reports earnings on October 29.',
  },
  {
    t: 'DLR', n: 'Digital Realty', s: 'Real Estate', build: true,
    cap: '$65.5B', fpe: 'n/m', div: '2.81%', beta: '1.01', earn: 'Oct 29', hi: 208.14, lo: 146.23,
    what: 'One of the world’s largest data-center landlords. Digital Realty develops and owns data centers and leases space, power and network connections to cloud and AI companies. It is organized as a real estate investment trust (REIT).',
    why: 'Digital Realty owns the finished buildings, which balances the contractors. Apollo’s chief economist argues that if the build-out slows, owners of scarce capacity win and builders lose. That means DLR should hold up in the scenario that hurts EME, PWR and STRL. In Q1 2026 it signed $707 million of new annual rent, including the largest hyperscale lease in its history, and it pays a 2.8% dividend.',
    risk: 'REITs fall when interest rates rise. Its P/E looks huge (about 83) only because REITs record large depreciation charges. They are normally valued on funds from operations (FFO) instead. It reports earnings on October 29.',
  },
  {
    t: 'MSFT', n: 'Microsoft', s: 'Information Technology', build: false,
    cap: '$3.93T', fpe: '26.8', div: '0.74%', beta: '1.10', earn: 'Oct 28', hi: 553.72, lo: 349.20,
    what: 'Software and cloud: Windows, Microsoft 365, the Azure cloud, LinkedIn and Xbox.',
    why: 'Most of Microsoft’s revenue comes from recurring subscriptions, which stay steady in a slow economy. It is also one of the biggest customers for new data centers, so it represents the side of the theme that pays for them. In the October selloff, the Big Tech companies paying for the build-out fell about 5%, while the builders fell a median of 23%.',
    risk: 'The stock trades close to its 52-week high, so a disappointing earnings report on October 28 has room to hurt it.',
  },
  {
    t: 'GOOGL', n: 'Alphabet', s: 'Communication Services', build: false,
    cap: '$3.92T', fpe: '23.9', div: '0.27%', beta: '1.21', earn: 'Oct 28', hi: 408.61, lo: 264.28,
    what: 'Google Search, YouTube, Android, Google Cloud, and Waymo self-driving cars.',
    why: 'Its ad business produces enormous cash flow, and it trades at a lower valuation than most mega-cap tech stocks. Alphabet raised its 2026 capital spending plan to $195–205 billion, and much of that money pays the kind of companies in the first half of this portfolio.',
    risk: 'Antitrust cases, and AI chatbots competing with search. The stock is about 22% below its 52-week high. It reports earnings on October 28.',
  },
  {
    t: 'TJX', n: 'TJX Companies', s: 'Consumer Discretionary', build: false,
    cap: '$159.6B', fpe: '26.8', div: '1.32%', beta: '0.56', earn: 'Nov 18', hi: 170.00, lo: 122.78,
    what: 'The largest off-price retailer: T.J. Maxx, Marshalls, HomeGoods and Sierra in the U.S., Winners in Canada, and T.K. Maxx in Europe.',
    why: 'It does well in good economies and bad ones. When prices rise (September inflation was 3.7%), shoppers trade down to off-price stores. It has nothing to do with AI, so it is a true diversifier. A beta of 0.56 means it moves about half as much as the market.',
    risk: 'Tariffs on imported goods and a slowdown in consumer spending. At about 27 times earnings, it is not cheap.',
  },
  {
    t: 'PG', n: 'Procter & Gamble', s: 'Consumer Staples', build: false,
    cap: '$370.4B', fpe: '22.6', div: '2.73%', beta: '0.38', earn: 'Reported Oct 22', hi: 167.25, lo: 137.62,
    what: 'Household brands: Tide, Pampers, Gillette, Crest, Charmin, Bounty and Head & Shoulders.',
    why: 'People buy laundry detergent and toothpaste in any economy. With a beta of 0.38, it is one of the steadiest stocks in the market. It pays a 2.7% dividend and has raised it every year for more than 65 years.',
    risk: 'Growth is slow, and inflation raises its costs. It will not keep up in a strong bull market.',
  },
  {
    t: 'JNJ', n: 'Johnson & Johnson', s: 'Health Care', build: false,
    cap: '$640.3B', fpe: '23.9', div: '2.02%', beta: '0.21', earn: 'Reported Oct 13', hi: 281.07, lo: 184.66,
    what: 'Prescription drugs (cancer, immunology, neuroscience) and medical devices (surgical tools, orthopedics, heart devices).',
    why: 'Demand for health care does not depend on the economy or on AI. With a beta of 0.21, it barely follows the market. It has one of the strongest balance sheets in the country and pays a 2.0% dividend. The stock is near its 52-week high, which shows strong momentum.',
    risk: 'Talc lawsuits, Medicare drug-price negotiations, and competition as older drugs lose patent protection.',
  },
  {
    t: 'JPM', n: 'JPMorgan Chase', s: 'Financials', build: false,
    cap: '$888.3B', fpe: '13.3', div: '1.97%', beta: '1.02', earn: 'Reported Oct 13', hi: 366.50, lo: 279.10,
    what: 'The largest U.S. bank: Chase consumer banking, investment banking, and asset and wealth management.',
    why: 'Higher interest rates hurt the REIT and the utility in this portfolio, but they usually help banks earn more on loans. JPMorgan offsets that risk. At about 13 times earnings, it is the cheapest stock in the portfolio, and it pays a 2.0% dividend.',
    risk: 'A recession would increase loan losses.',
  },
  {
    t: 'NEE', n: 'NextEra Energy', s: 'Utilities', build: false,
    cap: '$172.4B', fpe: '19.9', div: '3.03%', beta: '0.64', earn: 'Reported Oct 21', hi: 98.75, lo: 74.41,
    what: 'Owner of Florida Power & Light, one of the largest regulated utilities, and NextEra Energy Resources, one of the world’s largest wind and solar developers.',
    why: 'A regulated utility earns steady, predictable profits and pays the highest dividend in this portfolio (3.0%). Its development arm builds the power plants that data centers need, including a deal with Google to restart Iowa’s Duane Arnold nuclear plant. That gives some exposure to the theme without the volatility of a contractor.',
    risk: 'Utilities fall when bond yields rise, and changes to clean-energy tax credits could hurt its renewables business.',
  },
  {
    t: 'XOM', n: 'Exxon Mobil', s: 'Energy', build: false,
    cap: '$755.0B', fpe: '12.6', div: '2.24%', beta: '0.21', earn: 'Oct 30', hi: 184.37, lo: 112.51,
    what: 'The largest U.S. oil and gas company: drilling, refining and chemicals.',
    why: 'Exxon hedges the thing that is hurting stocks right now. Gasoline drove more than half of September’s inflation jump. When oil spikes, tech and construction stocks usually fall, while Exxon earns more. It is cheap at about 12.6 times next year’s earnings and pays a 2.2% dividend.',
    risk: 'It is trading at its 52-week high, so if oil prices fall back, it falls too. It reports earnings on October 30.',
  },
];

const portfolio = new Map(fs.readFileSync(path.join(ROOT, 'portfolio.csv'), 'utf8').trim().split(/\r?\n/).slice(1)
  .map((line) => line.split(',')).map(([symbol, company, sector, shares, price]) => [symbol, { company, sector, shares, price }]));
if (portfolio.size !== H.length) throw new Error('portfolio.csv and the write-up list different companies');
H.forEach((h) => {
  const row = portfolio.get(h.t);
  if (!row || row.company !== h.n || row.sector !== h.s) throw new Error(`${h.t} does not match portfolio.csv`);
  h.p = Number(row.price);
  h.sh = Number(row.shares);
  h.costC = Math.round(h.p * 100) * h.sh;
});
const totalC = H.reduce((a, h) => a + h.costC, 0);
const buildC = H.filter((h) => h.build).reduce((a, h) => a + h.costC, 0);
const balanceC = totalC - buildC;
const money = (c) => '$' + (c / 100).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const price = (p) => '$' + p.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const pct = (c) => (100 * c / totalC).toFixed(1) + '%';
const sectorOrder = ['Industrials', 'Materials', 'Real Estate', 'Information Technology', 'Communication Services',
  'Consumer Discretionary', 'Consumer Staples', 'Health Care', 'Financials', 'Utilities', 'Energy'];
const sectors = sectorOrder.map((s) => {
  const hs = H.filter((h) => h.s === s);
  return { s, hs, c: hs.reduce((a, h) => a + h.costC, 0) };
});
if (sectors.some((x) => x.hs.length === 0)) throw new Error('sector with no holding');
if (totalC < BUDGET_MIN * 100 || totalC > BUDGET_MAX * 100) throw new Error('outside budget range');
const roundPct = (c) => Math.round(100 * c / totalC) + '%';
const industrialsC = sectors.find((x) => x.s === 'Industrials').c;

// ---- styling ----
const FONT = 'Calibri';
const INK = '1F2937';
const MUTED = '6B7280';
const ACCENT = '0F4C5C';
const BUILD_FILL = 'E3F0F2';
const HEAD_FILL = '0F4C5C';
const ZEBRA = 'F5F7F8';
const RULE = 'D1D5DB';
const CONTENT_W = 9360;

const border = { style: BorderStyle.SINGLE, size: 4, color: RULE };
const borders = { top: border, bottom: border, left: border, right: border };

const run = (text, o = {}) => new TextRun({ text, font: FONT, color: INK, ...o });
const para = (children, o = {}) => new Paragraph({
  children: Array.isArray(children) ? children : [run(children)],
  spacing: { after: 120, line: 276 }, ...o,
});
const h1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text })] });
const h2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text })] });
const bullet = (children) => new Paragraph({
  numbering: { reference: 'bullets', level: 0 },
  children: Array.isArray(children) ? children : [run(children)],
  spacing: { after: 80, line: 276 },
});
const labelled = (label, text) => para([run(label + ' ', { bold: true, color: ACCENT }), run(text)]);

function cell(text, width, o = {}) {
  const { bold = false, fill, align = AlignmentType.LEFT, color = INK, size = 19 } = o;
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ alignment: align, children: [new TextRun({ text, bold, color, font: FONT, size })] })],
  });
}

function table(widths, header, rows, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  if (total !== CONTENT_W) throw new Error('table widths sum to ' + total);
  const aligns = o.aligns || widths.map(() => AlignmentType.LEFT);
  const head = new TableRow({
    tableHeader: true,
    children: header.map((h, i) => cell(h, widths[i], { bold: true, fill: HEAD_FILL, color: 'FFFFFF', align: aligns[i] })),
  });
  const body = rows.map((r, ri) => {
    if (r.group) {
      return new TableRow({ children: [new TableCell({
        borders, columnSpan: widths.length, width: { size: CONTENT_W, type: WidthType.DXA },
        shading: { fill: r.fill, type: ShadingType.CLEAR, color: 'auto' },
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [new Paragraph({ children: [new TextRun({ text: r.group, bold: true, color: ACCENT, font: FONT, size: 19 })] })],
      })] });
    }
    return new TableRow({
      children: r.cells.map((v, i) => cell(v, widths[i], {
        align: aligns[i], bold: r.bold, fill: r.fill || (ri % 2 ? ZEBRA : undefined),
      })),
    });
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    rows: [head, ...body],
  });
}

const R = AlignmentType.RIGHT;
const L = AlignmentType.LEFT;
const C = AlignmentType.CENTER;
const spacer = () => new Paragraph({ children: [], spacing: { after: 60 } });

// ---- content ----
const children = [];

children.push(new Paragraph({
  children: [new TextRun({ text: 'Stock Picks: Data-Center Build-Out Tilt', bold: true, size: 44, color: ACCENT, font: FONT })],
  spacing: { after: 80 },
}));
children.push(new Paragraph({
  children: [run('13 companies across all 11 sectors, with extra weight on the companies that physically build data centers', { size: 24, color: MUTED })],
  spacing: { after: 80 },
}));
children.push(new Paragraph({
  children: [run('Prepared October 26, 2026. Prices are live quotes from 11:30–11:34 AM EDT.', { size: 20, color: MUTED, italics: true })],
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 6 } },
  spacing: { after: 240 },
}));

children.push(h1('At a glance'));
children.push(bullet([run('Budget: ', { bold: true }), run('$20,000 to $25,000')]));
children.push(bullet([run('Invested: ', { bold: true }), run(`${money(totalC)} in 13 companies, inside the budget range.`)]));
children.push(bullet([run('Data-center build-out: ', { bold: true }), run(`${money(buildC)} (${pct(buildC)}) in five companies that build, supply or own data centers. The other eight companies get about 5–6% each.`)]));
children.push(bullet([run('Sectors: ', { bold: true }), run('11 of 11. Every sector in the stock market is represented.')]));

children.push(h2('Rules check'));
children.push(table([3600, 1800, 2460, 1500], ['Requirement', 'Required', 'This portfolio', 'Met?'], [
  { cells: ['Number of companies', '10 or more', '13', 'Yes'] },
  { cells: ['Sectors represented', '6 of 11', '11 of 11', 'Yes'] },
  { cells: ['Total spend', '$20,000–$25,000', money(totalC), 'Yes'] },
], { aligns: [L, C, C, C] }));
children.push(spacer());

children.push(h1('The plan'));
children.push(para('The portfolio has two halves with two different jobs.'));
children.push(labelled(`The builders (about ${roundPct(buildC)}).`, 'These five companies make money when a data center goes up. Sterling prepares the site. Vulcan supplies the stone for the concrete. Quanta connects the site to the power grid. EMCOR installs the electrical and cooling systems. Digital Realty owns and leases out the finished building. They are spread across three sectors (Industrials, Materials and Real Estate), and they include an owner as well as contractors. That matters because owners and builders react differently if the boom slows.'));
children.push(labelled(`The balance (about ${roundPct(balanceC)}).`, 'One company from each of the other eight sectors. Each was chosen because it holds up when the builders struggle, for a specific reason: steady subscriptions (Microsoft), recession-proof demand (P&G, Johnson & Johnson), gains from higher interest rates (JPMorgan), gains from higher oil prices (Exxon), or shoppers trading down (TJX). Most of them move much less than the market, with a beta well under 1.'));

children.push(h1('Holdings'));
const hw = [900, 2200, 2300, 1000, 800, 1160, 1000];
const holdingRow = (h) => ({
  cells: [h.t, h.n, h.s, price(h.p), String(h.sh), money(h.costC), pct(h.costC)],
  fill: h.build ? BUILD_FILL : 'FFFFFF',
});
children.push(table(hw, ['Ticker', 'Company', 'Sector', 'Price', 'Shares', 'Cost', 'Weight'], [
  { group: `The builders: ${money(buildC)} (${pct(buildC)})`, fill: 'C9E2E6' },
  ...H.filter((h) => h.build).map(holdingRow),
  { group: `The balance: ${money(balanceC)} (${pct(balanceC)})`, fill: 'E5E7EB' },
  ...H.filter((h) => !h.build).map(holdingRow),
  { cells: ['', 'Total invested', '', '', '', money(totalC), '100%'], bold: true, fill: 'E5E7EB' },
], { aligns: [L, L, L, R, R, R, R] }));
children.push(para([run('Weights are a share of the amount invested.', { size: 18, color: MUTED, italics: true })], { spacing: { before: 80, after: 200 } }));

children.push(h2('Sector breakdown'));
children.push(table([2900, 3460, 1700, 1300], ['Sector', 'Companies', 'Cost', 'Weight'],
  sectors.map((x) => ({
    cells: [x.s, x.hs.map((h) => h.n).join(', '), money(x.c), pct(x.c)],
    fill: x.hs.some((h) => h.build) ? BUILD_FILL : undefined,
  })),
  { aligns: [L, L, R, R] }));
children.push(spacer());

children.push(h1('Before you buy: what is happening right now'));
children.push(para('The data-center builders are in a sharp selloff, and that affects when to buy:'));
children.push(bullet([run('Since the S&P 500’s record close on October 6, the companies that build and equip data centers are down a median of ', {}), run('23%', { bold: true }), run('. Power producers are down 14%, Nvidia 11%, and the Big Tech companies paying for the build-out only about 5%.')]));
children.push(bullet('The business numbers have not cracked. EMCOR’s backlog is at a record. Comfort Systems’ backlog nearly doubled. The four biggest tech companies have raised their combined 2026 capital spending plans to $720–745 billion. The selloff is about investor nerves and hot inflation (3.7% in September), not about canceled projects.'));
children.push(bullet([run('This week is very busy. ', { bold: true }), run('Sterling, Microsoft and Alphabet report on October 28, the same day as the Federal Reserve’s rate decision. EMCOR, Quanta, Vulcan and Digital Realty report on October 29, and Exxon on October 30.')]));
children.push(labelled('What this means:', 'buying now gets the builders at a big discount, but it is not a bottom signal. Falling stocks can keep falling. If the class buy date is flexible, waiting until after October 30 means buying after the earnings reports and capital-spending updates instead of before them. Prices could be higher or lower by then. If the buy date is fixed, the plan still works, because the balance half is built for exactly this kind of week.'));

children.push(h1('The 13 companies'));
H.forEach((h, i) => {
  children.push(h2(`${i + 1}. ${h.n} (${h.t})`));
  children.push(para([
    run(`${h.s}  ·  ${h.build ? 'Builder' : 'Balance'}  ·  ${h.sh} shares × ${price(h.p)} = ${money(h.costC)} (${pct(h.costC)})`, { bold: true, color: MUTED, size: 20 }),
  ], { spacing: { after: 100 } }));
  children.push(labelled('What it does:', h.what));
  children.push(labelled('Why it’s here:', h.why));
  children.push(labelled('Main risk:', h.risk));
  children.push(para([run(
    `Market cap ${h.cap}  ·  Forward P/E ${h.fpe}  ·  Dividend ${h.div}  ·  Beta ${h.beta}  ·  52-week range ${price(h.lo)}–${price(h.hi)}  ·  Earnings ${h.earn}`,
    { size: 18, color: MUTED },
  )], { spacing: { after: 200 } }));
});

children.push(h1('Risks for the whole portfolio'));
children.push(bullet([run('Theme concentration. ', { bold: true }), run(`About ${roundPct(buildC)} depends on data-center construction continuing. Microsoft, Alphabet and NextEra are also connected to it. If tech companies cut spending sharply, more than half the portfolio would feel it.`)]));
children.push(bullet([run(`Industrials is ${roundPct(industrialsC)} on its own. `, { bold: true }), run('That is the price of the tilt you asked for. EMCOR alone is about 16%. The three contractors often move together, as they did on October 22, when all three fell 5–9% in one day.')]));
children.push(bullet([run('Interest rates. ', { bold: true }), run('Digital Realty and NextEra fall when rates rise. JPMorgan is the offset.')]));
children.push(bullet([run('Nobody can time this. ', { bold: true }), run('The builders could keep falling after the buy. The portfolio is designed to survive that, not to avoid it.')]));

children.push(h1('Swaps considered'));
children.push(para('These were close calls and can be swapped in if the class wants a change:'));
children.push(table([2700, 6660], ['Company', 'Why it was left out'], [
  { cells: ['Comfort Systems USA (FIX)', 'A great data-center HVAC contractor (backlog of $17.8 billion), but it costs $1,257.74 a share. One share would be 6% of the budget, which makes the sizing too lumpy.'] },
  { cells: ['Vertiv (VRT)', 'Makes power and cooling equipment for data centers. It is an equipment maker, not a builder, and it would be a fourth Industrials stock.'] },
  { cells: ['Freeport-McMoRan (FCX)', 'Copper for data-center wiring. Its price swings with commodity prices and China. Could replace Vulcan.'] },
  { cells: ['Williams Companies (WMB)', 'Gas pipelines that also build power plants for data centers. Could replace Exxon for more theme exposure, but it would lose the oil-price hedge.'] },
  { cells: ['Amazon (AMZN)', 'The biggest spender on data centers. Could replace TJX, but it would add more AI exposure where the plan needs less.'] },
  { cells: ['Constellation (CEG) / Vistra (VST)', 'Power producers that sell to data centers. They are more tied to the AI trade than NextEra, and they fell 14% in the selloff.'] },
], { aligns: [L, L] }));
children.push(spacer());

children.push(h1('Sources'));
const src = (label, url) => bullet([
  new ExternalHyperlink({ link: url, children: [new TextRun({ text: label, style: 'Hyperlink', font: FONT })] }),
]);
children.push(bullet('Prices, 52-week ranges, market cap, P/E, dividend, beta and earnings dates: Stock Analysis (stockanalysis.com), October 26, 2026.'));
children.push(src('Yahoo Finance (Oct 23, 2026): The AI sell-off is crushing stocks of companies leading the data center build-out', 'https://finance.yahoo.com/markets/article/the-ai-sell-off-is-crushing-stocks-of-companies-leading-the-data-center-build-out-175206076.html'));
children.push(src('EMCOR Group: Second Quarter 2026 Results (SEC 8-K)', 'https://www.sec.gov/Archives/edgar/data/0000105634/000010563426000112/eme-ex991_2026630xq2.htm'));
children.push(src('Digital Realty: First Quarter 2026 Results (SEC 8-K)', 'https://www.sec.gov/Archives/edgar/data/0001297996/000110465926047702/dlr-20260423xex99d1.htm'));
children.push(src('TMT Finance (Aug 18, 2026): 2026 hyperscaler capex tops US$700bn', 'https://www.tmtfinance.com/intel/2026-hyperscaler-capex-tops-us700bn-analysis'));
children.push(src('AlphaStreet: Sterling Infrastructure drops 9.4% amid sector-wide selling', 'https://news.alphastreet.com/sterling-infrastructure-drops-9-4-amid-sector-wide-selling/'));
children.push(src('Forex Fundamentals: U.S. CPI, September 2026 (released Oct 14, 2026)', 'https://forexfundamentals.com/economic-calendar/us-cpi/2026-10-14'));
children.push(src('StockTitan: Earnings calendar for the week of October 26, 2026', 'https://www.stocktitan.net/articles/earnings-calendar-week-of-october-26-2026'));
children.push(para([run('This is a class project, not financial advice. Prices move every second, so recheck them on the buy date and make sure the total stays between $20,000 and $25,000.', { size: 18, color: MUTED, italics: true })], { spacing: { before: 160 } }));

const doc = new Document({
  creator: 'Stock project',
  title: 'Stock Picks: Data-Center Build-Out Tilt',
  styles: {
    default: { document: { run: { font: FONT, size: 22, color: INK } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 30, bold: true, color: ACCENT, font: FONT },
        paragraph: { spacing: { before: 320, after: 140 }, outlineLevel: 0, keepNext: true } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 25, bold: true, color: INK, font: FONT },
        paragraph: { spacing: { before: 220, after: 60 }, outlineLevel: 1, keepNext: true } },
    ],
  },
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 540, hanging: 270 } } } }],
    }],
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: 'Page ', size: 16, color: MUTED, font: FONT }),
            new TextRun({ children: [PageNumber.CURRENT], size: 16, color: MUTED, font: FONT })],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log(`wrote ${OUT}`);
  console.log(`invested ${money(totalC)}  builders ${money(buildC)} (${pct(buildC)})`);
  sectors.forEach((x) => console.log(`  ${x.s.padEnd(24)} ${money(x.c).padStart(10)} ${pct(x.c)}`));
});
