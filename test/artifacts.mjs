// Runs the app, saves the workbook and highlighted PDFs it produced to test/out/,
// then exercises the review round-trip: write a decision into the Review sheet's
// COMMENTS column, feed it back, and confirm the summary moves.
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
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

const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
const page = await browser.newPage();
page.on('pageerror', (e) => console.log('[pageerror]', e.message));

const which = process.argv[2] || 'amex';
const FILES = which === 'cibc'
  ? ['samples/CIBC Statement - Jul 13 2026.pdf', 'samples/CIBC Statement - Aug 14 2026.pdf']
  : ['samples/Amex Jul 17 2026 Statement.pdf', 'samples/Amex Aug 17 2026 Statement.pdf'];

await page.goto(`${base}/index.html`, { waitUntil: 'domcontentloaded' });
await page.waitForFunction(() => !!window.ExcelJS && !!window.PDFLib, null, { timeout: 30000 });
await page.setInputFiles('#fileInput', FILES.map((f) => path.join(ROOT, f)));
await page.click('#runBtn');
await page.waitForSelector('#result:not([hidden])', { timeout: 180000 });

// ------------------------------------------------------ save what it built
const artifacts = await page.evaluate(async () => {
  const r = window.__pipelineResult;
  const b64 = (u8) => {
    let s = '';
    for (let i = 0; i < u8.length; i += 8192) s += String.fromCharCode(...u8.subarray(i, i + 8192));
    return btoa(s);
  };
  return {
    workbook: { name: r.workbook.name, data: b64(r.workbook.bytes) },
    pdfs: r.categoryPdfs.map((p) => ({ name: p.fileName, lines: p.lines, data: b64(p.bytes) })),
  };
});

fs.writeFileSync(path.join(OUT, artifacts.workbook.name), Buffer.from(artifacts.workbook.data, 'base64'));
console.log(`saved ${artifacts.workbook.name}`);
for (const p of artifacts.pdfs) {
  fs.writeFileSync(path.join(OUT, p.name), Buffer.from(p.data, 'base64'));
  console.log(`saved ${p.name} (${p.lines} highlights)`);
}

// ------------------------------------------------------ review round-trip
const reviewed = process.argv[3];
if (reviewed) {
  console.log(`\napplying review from ${reviewed}`);
  await page.click('.tab[data-tab="review"]');
  await page.setInputFiles('#reviewInput', path.resolve(reviewed));
  await page.waitForSelector('#reviewResult:not([hidden])', { timeout: 180000 });
  const after = await page.evaluate(() => ({
    applied: window.__pipelineResult.applied,
    categories: window.__lastResult ? null : null,
    summary: window.__pipelineResult.summary,
    review: window.__pipelineResult.review.length,
  }));
  console.log('  applied:', after.applied, '| still in review:', after.review);
  console.log('  summary:', JSON.stringify(after.summary));
}

await browser.close();
server.close();
