// Dump what the Amex/CIBC parsers see, so a mis-parse can be traced to a row.
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
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

const target = process.argv[2] || 'samples/Amex Jul 17 2026 Statement.pdf';
const pageNo = Number(process.argv[3] ?? 1);

const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
const page = await browser.newPage();
page.on('pageerror', (e) => console.log('[pageerror]', e.message));
await page.goto(`${base}/index.html`, { waitUntil: 'domcontentloaded' });

const out = await page.evaluate(async (args) => {
  const { url, pageNo } = args;
  const pdfjsLib = await import('/vendor/pdf.min.mjs');
  pdfjsLib.GlobalWorkerOptions.workerSrc = args.worker;
  const base = await import('/js/parsers/base.js');
  const bytes = await (await fetch(url)).arrayBuffer();
  const { pages } = await base.readPages(pdfjsLib, bytes);
  const rows = base.rowsOf(pages[pageNo]);
  return {
    pageCount: pages.length,
    rows: rows.map((r) => ({ text: r.text, y0: Math.round(r.y0), n: r.words.length })),
  };
}, { url: `${base}/${encodeURI(target)}`, pageNo, worker: `${base}/vendor/pdf.worker.min.mjs` });

console.log(`${target} — ${out.pageCount} pages; page ${pageNo + 1} rows:`);
out.rows.forEach((r, i) => console.log(`${String(i).padStart(3)} y=${r.y0} n=${r.n}  ${r.text.slice(0, 120)}`));

await browser.close();
server.close();
