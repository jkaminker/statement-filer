// Drives the real app in a headless browser against the actual statements and
// checks the parse ties to the statement control totals.
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const MIME = {
  '.html': 'text/html', '.css': 'text/css', '.js': 'text/javascript',
  '.mjs': 'text/javascript', '.json': 'application/json', '.pdf': 'application/pdf',
};

const server = http.createServer((req, res) => {
  const file = path.join(ROOT, decodeURIComponent(req.url.split('?')[0]).replace(/^\//, '') || 'index.html');
  if (!file.startsWith(ROOT) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
    res.writeHead(404).end('nope');
    return;
  }
  res.writeHead(200, { 'Content-Type': MIME[path.extname(file)] || 'application/octet-stream' });
  fs.createReadStream(file).pipe(res);
});
await new Promise((r) => server.listen(0, r));
const base = `http://127.0.0.1:${server.address().port}`;

// Expected category totals are real figures, so this file is gitignored. Without it the
// suite still checks the parse and the reconciliation, just not the per-category split.
const EXP_PATH = path.join(ROOT, 'test/expected.json');
const EXPECTED = fs.existsSync(EXP_PATH) ? JSON.parse(fs.readFileSync(EXP_PATH, 'utf8')) : {};

const CASES = [
  {
    name: 'Amex Q3 2026',
    files: ['samples/Amex Jul 17 2026 Statement.pdf', 'samples/Amex Aug 17 2026 Statement.pdf'],
    expect: { card: 'amex', quarter: 'Q3 2026', total: 46035.66, count: 162,
              categories: EXPECTED.amex },
  },
  {
    name: 'CIBC Q3 2026',
    files: ['samples/CIBC Statement - Jul 13 2026.pdf', 'samples/CIBC Statement - Aug 14 2026.pdf'],
    expect: { card: 'cibc', quarter: 'Q3 2026', total: 6555.44, count: 91,
              categories: EXPECTED.cibc },
  },
];

const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
const page = await browser.newPage();
page.on('console', (m) => { if (m.type() === 'error') console.log('  [console]', m.text()); });
page.on('pageerror', (e) => console.log('  [pageerror]', e.message));

let failures = 0;
for (const c of CASES) {
  console.log(`\n=== ${c.name}`);
  await page.goto(base, { waitUntil: 'domcontentloaded' });
  // mirror what the user does once per browser: import the private merchant map
  if (fs.existsSync(path.join(ROOT, 'my-rules.json'))) {
    const priv = fs.readFileSync(path.join(ROOT, 'my-rules.json'), 'utf8');
    await page.evaluate((json) => localStorage.setItem('statement-filer.rules', json), priv);
    await page.reload({ waitUntil: 'domcontentloaded' });
  }
  await page.waitForFunction(() => !!window.ExcelJS && !!window.PDFLib, null, { timeout: 30000 });
  await page.setInputFiles('#fileInput', c.files.map((f) => path.join(ROOT, f)));
  await page.click('#runBtn');
  await page.waitForSelector('#result:not([hidden])', { timeout: 120000 }).catch(() => {});

  const out = await page.evaluate(() => {
    const log = document.getElementById('log').innerText;
    const res = document.getElementById('result');
    return { log, visible: !res.hidden, text: res.innerText };
  });

  if (!out.visible) {
    console.log('  FAILED — no result rendered');
    console.log(out.log.split('\n').map((l) => '    ' + l).join('\n'));
    failures++;
    continue;
  }

  const data = await page.evaluate(() => window.__lastResult || null);
  console.log(out.text.split('\n').slice(0, 4).map((l) => '  ' + l).join('\n'));

  const got = data || {};
  const checks = [
    ['card', got.card, c.expect.card],
    ['quarter', got.quarter, c.expect.quarter],
    ['transactions', got.count, c.expect.count],
    ['total', got.total, c.expect.total],
    ['variance', got.variance, 0],
  ];
  for (const [label, actual, want] of checks) {
    const ok = typeof want === 'number' ? Math.abs((actual ?? NaN) - want) < 0.005 : actual === want;
    console.log(`  ${ok ? 'ok  ' : 'FAIL'} ${label}: ${actual}${ok ? '' : ` (expected ${want})`}`);
    if (!ok) failures++;
  }
  if (c.expect.categories) {
    const want = c.expect.categories;
    const keys = [...new Set([...Object.keys(want), ...Object.keys(got.categories || {})])].sort();
    for (const k of keys) {
      const a = (got.categories || {})[k], w = want[k];
      const ok = Math.abs((a ?? 0) - (w ?? 0)) < 0.005;
      console.log(`  ${ok ? 'ok  ' : 'FAIL'} ${k}: ${a ?? '-'}${ok ? '' : ` (expected ${w ?? '-'})`}`);
      if (!ok) failures++;
    }
  }
  console.log(`  review rows: ${got.review}  |  category PDFs: ${got.pdfs}`);
}

await browser.close();
server.close();
console.log(failures ? `\n${failures} check(s) failed` : '\nAll checks passed');
process.exit(failures ? 1 : 0);
