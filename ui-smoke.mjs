// Boots the real page in a headless browser and checks the new controls exist,
// wire up, and don't throw. No statements needed — this is about the shell.
//
//   node test/ui-smoke.mjs

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
  const f = path.join(ROOT, decodeURIComponent(req.url.split('?')[0]).replace(/^\//, '') || 'index.html');
  if (!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) {
    res.writeHead(404).end('nope');
    return;
  }
  res.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' });
  fs.createReadStream(f).pipe(res);
});
await new Promise((r) => server.listen(0, r));
const base = `http://127.0.0.1:${server.address().port}`;

let failures = 0;
const errors = [];
const check = (ok, label, detail = '') => {
  console.log(`  ${ok ? 'ok  ' : 'FAIL'} ${label}${detail ? `  — ${detail}` : ''}`);
  if (!ok) failures++;
};

const browser = await chromium.launch({
  executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
});
const page = await browser.newPage();
// Google's sign-in script is a third-party <script> and this sandbox has no
// route to it. That is an environment fact, not a defect in the page, so it is
// the one failure this smoke test forgives.
const isExternal = (t) => /Failed to load resource/i.test(t) && !t.includes('127.0.0.1');
page.on('pageerror', (e) => errors.push(e.message));
page.on('console', (m) => {
  if (m.type() === 'error' && !isExternal(m.text())) errors.push(m.text());
});

await page.goto(`${base}/index.html`, { waitUntil: 'domcontentloaded' });
await page.waitForFunction(() => !!window.ExcelJS && !!window.PDFLib, null, { timeout: 30000 });
await page.waitForTimeout(400);

console.log('--- A. the page boots clean');
check(errors.length === 0, 'no uncaught errors on load', errors.join(' | '));

console.log('\n--- B. the new controls are there');
check(await page.locator('#preflight').count() === 1, 'preflight panel exists');
check(await page.locator('#preflight').isHidden(), 'and is hidden until files are picked');
check(await page.locator('input[name="runMode"]').count() === 2, 'both run modes exist');
check(
  await page.locator('input[name="runMode"][value="append"]').isChecked(),
  'append is the default'
);
check(await page.locator('#pfQuarter').count() === 1, 'the quarter override exists');

console.log('\n--- C. the What\'s filed tab');
await page.click('.tab[data-tab="filed"]');
await page.waitForTimeout(200);
check(await page.locator('#tab-filed').isVisible(), 'the tab opens');
const cards = await page.locator('#filedCard option').allTextContents();
check(cards.length === 3, 'every card is offered', cards.join(', '));
const fys = await page.locator('#filedFy option').allTextContents();
check(fys.some((t) => /Annual Audit \d{4} Sep/.test(t)), 'audit years read from the rules',
  fys.join(', '));
check(
  await page.locator('#filedLoadBtn').isDisabled(),
  'Open is disabled while Drive is not connected'
);

console.log('\n--- D. the run button says what it will do');
await page.click('.tab[data-tab="process"]');
check(
  (await page.locator('#runBtn').textContent()).trim() === 'Analyze',
  'it reads "Analyze" with nothing loaded'
);

// drive it the way main.js does once a preflight has landed, without a PDF
await page.evaluate(() => {
  document.querySelector('input[name="runMode"][value="rebuild"]').checked = true;
  document.querySelector('input[name="runMode"][value="rebuild"]')
    .dispatchEvent(new Event('change', { bubbles: true }));
});
check(
  (await page.locator('#runBtn').textContent()).trim() === 'Analyze',
  'and stays neutral until a preflight has actually run'
);

console.log('\n--- E. still no errors after poking at it');
check(errors.length === 0, 'clean console throughout', errors.join(' | '));

await browser.close();
server.close();
console.log(failures ? `\n${failures} check(s) failed.` : '\nAll checks passed.');
process.exit(failures ? 1 : 0);
