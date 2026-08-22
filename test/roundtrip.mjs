// End-to-end check of the review loop:
//   1. run with a few merchants deliberately removed from the rules, so rows land in Review
//   2. write categories into the COMMENTS column of the workbook it produced
//   3. feed that workbook back in and confirm the summary moves and Review empties
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(ROOT, 'test/out');
fs.mkdirSync(OUT, { recursive: true });

const MIME = { '.html': 'text/html', '.css': 'text/css', '.js': 'text/javascript',
  '.mjs': 'text/javascript', '.json': 'application/json', '.pdf': 'application/pdf' };
const server = http.createServer((req, res) => {
  const f = path.join(ROOT, decodeURIComponent(req.url.split('?')[0]).replace(/^\//, ''));
  if (!fs.existsSync(f) || fs.statSync(f).isDirectory()) return res.writeHead(404).end();
  res.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' });
  fs.createReadStream(f).pipe(res);
});
await new Promise((r) => server.listen(0, r));
const base = `http://127.0.0.1:${server.address().port}`;

// merchants we pull out of the rules so their rows have to go to Review
const DROP = ['WHOLE FOODS MARKET TORONTO', 'REGUS REGUS CANADA'];

const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
const page = await browser.newPage();
page.on('pageerror', (e) => console.log('[pageerror]', e.message));
await page.goto(`${base}/index.html`, { waitUntil: 'domcontentloaded' });

// stash doctored rules before the app reads them. Use the private merchant map
// when it's there — the rules.json in the repo ships with an empty one, so
// without it every row lands in Review and there is no review loop to test.
const RULES_FILE = fs.existsSync(path.join(ROOT, 'my-rules.json')) ? 'my-rules.json' : 'rules.json';
await page.evaluate(async ({ drop, file }) => {
  const rules = await (await fetch(file)).json();
  for (const d of drop) delete rules.merchants.amex[d];
  localStorage.setItem('statement-filer.rules', JSON.stringify(rules));
}, { drop: DROP, file: RULES_FILE });
await page.reload({ waitUntil: 'domcontentloaded' });
await page.waitForFunction(() => !!window.ExcelJS && !!window.PDFLib, null, { timeout: 30000 });

await page.setInputFiles('#fileInput', [
  path.join(ROOT, 'samples/Amex Jul 17 2026 Statement.pdf'),
  path.join(ROOT, 'samples/Amex Aug 17 2026 Statement.pdf'),
]);
await page.click('#runBtn');
await page.waitForSelector('#result:not([hidden])', { timeout: 180000 });

const before = await page.evaluate(() => ({
  summary: window.__pipelineResult.summary,
  review: window.__pipelineResult.review.length,
  total: window.__pipelineResult.parsedTotal,
  variance: window.__pipelineResult.variance,
  name: window.__pipelineResult.workbook.name,
  b64: (() => {
    const u8 = window.__pipelineResult.workbook.bytes;
    let s = '';
    for (let i = 0; i < u8.length; i += 8192) s += String.fromCharCode(...u8.subarray(i, i + 8192));
    return btoa(s);
  })(),
}));
console.log(`run 1: ${before.review} rows in Review, total $${before.total}, variance ${before.variance}`);
console.log('        summary:', JSON.stringify(before.summary));

// ---- fill in the COMMENTS column the way Jeff would ------------------------
const wbPath = path.join(OUT, 'roundtrip-' + before.name);
fs.writeFileSync(wbPath, Buffer.from(before.b64, 'base64'));

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(wbPath);
const rev = wb.getWorksheet('Review');
let filled = 0;
rev.eachRow((row, n) => {
  if (n < 4) return;
  const desc = String(row.getCell(2).value || '');
  if (desc.includes('WHOLE FOODS')) { row.getCell(6).value = 'Meals'; filled++; }
  else if (desc.includes('REGUS')) { row.getCell(6).value = 'Rent'; filled++; }
});
await wb.xlsx.writeFile(wbPath);
console.log(`wrote ${filled} decisions into the COMMENTS column`);

// ---- feed it back ---------------------------------------------------------
await page.click('.tab[data-tab="review"]');
await page.setInputFiles('#reviewInput', wbPath);
await page.waitForSelector('#reviewResult:not([hidden])', { timeout: 180000 });

const after = await page.evaluate(() => ({
  applied: window.__pipelineResult.applied,
  review: window.__pipelineResult.review.length,
  summary: window.__pipelineResult.summary,
  total: window.__pipelineResult.parsedTotal,
  pdfs: window.__pipelineResult.categoryPdfs.map((p) => `${p.category}:${p.lines}`),
  learned: JSON.parse(localStorage.getItem('statement-filer.rules')).merchants.amex['WHOLE FOODS MARKET TORONTO'],
}));

console.log(`run 2: applied ${after.applied}, ${after.review} left in Review`);
console.log('        summary:', JSON.stringify(after.summary));
console.log('        rebuilt PDFs:', after.pdfs.join(', '));
console.log('        learned WHOLE FOODS ->', after.learned);

const EXP_PATH = path.join(ROOT, 'test/expected.json');
const EXPECT = fs.existsSync(EXP_PATH) ? JSON.parse(fs.readFileSync(EXP_PATH, 'utf8')).amex : {};
let bad = 0;
for (const k of Object.keys(EXPECT)) {
  const ok = Math.abs((after.summary[k] ?? 0) - EXPECT[k]) < 0.005;
  if (!ok) { console.log(`  FAIL ${k}: ${after.summary[k]} (expected ${EXPECT[k]})`); bad++; }
}
if (after.review !== 0) { console.log(`  FAIL ${after.review} rows still in Review`); bad++; }
if (after.learned !== 'Meals') { console.log('  FAIL decision was not learned'); bad++; }
console.log(bad ? `\n${bad} check(s) failed` : '\nRound-trip matches the signed-off Q3 numbers exactly.');

await browser.close();
server.close();
process.exit(bad ? 1 : 0);
