// The merge round trip, without needing a statement PDF.
//
// A quarter is built up over several drops, and every drop rebuilds the
// workbook from what it reads back out of the last one. So anything the
// workbook cannot say is lost on the next merge. These checks pin down the two
// things that used to be lost: a row still sitting in Review, and the
// foreign-currency detail behind a row.
//
//   node test/roundtrip-merge.mjs

import ExcelJS from 'exceljs';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { buildWorkbook, readFiledWorkbook } = await import(
  path.join(ROOT, 'js/workbook.js')
);
const { carryForwardReview } = await import(path.join(ROOT, 'js/pipeline.js'));

const rules = JSON.parse(fs.readFileSync(path.join(ROOT, 'rules.json'), 'utf8'));
const REVIEW = rules.reviewCategory;

let failures = 0;
const check = (ok, label, detail = '') => {
  console.log(`  ${ok ? 'ok  ' : 'FAIL'} ${label}${detail ? `  — ${detail}` : ''}`);
  if (!ok) failures++;
};

// ─────────────────────────────────────────────────── a filed July workbook ──
// One settled row, one foreign-currency row, and one row you never answered.
const july = [
  { date: '2026-07-03', desc: 'TIM HORTONS #3025', amount: 12.45, category: 'Meals' },
  {
    date: '2026-07-11', desc: 'ADOBE SYSTEMS', amount: 92.4,
    category: 'Professional Fees',
    fx: { amount: 67.99, currency: 'USD', rate: 1.3591 },
  },
  { date: '2026-07-19', desc: 'NORTHERN SUPPLY CO', amount: 418.22, category: REVIEW },
];
const julyReview = [{
  ...july[2],
  suggested: 'Office Equipment',
  place: 'BRACEBRIDGE',
  note: 'New merchant on this card. Bank category: Retail and Grocery.',
}];

async function bytesOf(opts) {
  const wb = await buildWorkbook(ExcelJS, opts);
  return wb.xlsx.writeBuffer();
}

console.log('--- A. build July, read it back');
const julyBytes = await bytesOf({
  card: 'amex', quarter: 'Q3 2026', transactions: july, review: julyReview,
  flags: [], statements: [{ label: 'Jul 17 2026', controlTotal: 523.07 }],
  rules, notes: [],
});
const back = await readFiledWorkbook(ExcelJS, julyBytes, 'amex', rules);

check(back.rows.length === 3, 'all three rows come back', `${back.rows.length}`);
check(
  back.statements.length === 1 && back.statements[0].label === 'Jul 17 2026',
  'the statement label survives the reconciliation block',
  JSON.stringify(back.statements)
);

const adobe = back.rows.find((r) => r.desc.includes('ADOBE'));
check(!!adobe?.fx, 'the FX row keeps its foreign-currency detail');
check(
  adobe?.fx?.currency === 'USD' && Math.abs(adobe.fx.amount - 67.99) < 0.005,
  'the foreign amount and currency are exact',
  JSON.stringify(adobe?.fx)
);
check(
  Math.abs((adobe?.fx?.rate ?? 0) - 1.3591) < 1e-6,
  'the exchange rate is exact',
  String(adobe?.fx?.rate)
);

const pending = back.rows.find((r) => r.category === REVIEW);
check(!!pending, 'the unanswered row is still coded Review');
check(
  (pending?.note || '').includes('BRACEBRIDGE'),
  'its note comes back, town and all',
  pending?.note
);
check(
  pending?.suggested === 'Office Equipment',
  'its suggested category comes back',
  pending?.suggested
);

// ───────────────────────────── the merge a second drop would actually do ──
console.log('\n--- B. August merges in; nothing from July is dropped');
const august = [
  { date: '2026-08-04', desc: 'PRESTO FARE', amount: 3.35, category: 'Travel' },
  {
    date: '2026-08-21', desc: 'DIGITALOCEAN', amount: 61.08,
    category: 'Professional Fees',
    fx: { amount: 44.5, currency: 'USD', rate: 1.3726 },
  },
];
// the real thing the pipeline uses, not a copy of it
const carried = carryForwardReview(back.rows, rules);
check(carried.length === 1, 'the pipeline carries the unanswered row forward',
  `${carried.length} carried`);
check((carried[0]?.note || '').includes('BRACEBRIDGE'),
  'and hands it back with its note intact', carried[0]?.note);

const merged = [...back.rows, ...august].sort((a, b) => a.date.localeCompare(b.date));
const q3Bytes = await bytesOf({
  card: 'amex', quarter: 'Q3 2026', transactions: merged, review: carried,
  flags: [],
  statements: [
    { label: 'Jul 17 2026', controlTotal: 523.07 },
    { label: 'Aug 17 2026', controlTotal: 64.43 },
  ],
  rules, notes: [],
});

// read the merged workbook the way a THIRD drop would
const after = await readFiledWorkbook(ExcelJS, q3Bytes, 'amex', rules);
check(after.rows.length === 5, 'five rows after the merge', `${after.rows.length}`);
check(
  after.rows.filter((r) => r.fx).length === 2,
  'both foreign-currency rows survive — July\'s was not dropped',
  `${after.rows.filter((r) => r.fx).length}`
);
check(
  after.statements.length === 2,
  'both statements are listed in the reconciliation block',
  after.statements.map((s) => s.label).join(' & ')
);

const stillPending = after.rows.find((r) => r.category === REVIEW);
check(!!stillPending, 'July\'s unanswered row is still in Review after the merge');
check(
  (stillPending?.note || '').includes('BRACEBRIDGE'),
  'and it still carries its note, so it is answerable',
  stillPending?.note
);

// the Review SHEET, not just the Data sheet — this is the bug that hid items
const wb = new ExcelJS.Workbook();
await wb.xlsx.load(q3Bytes);
const revSheet = wb.getWorksheet(rules.cards.amex.sheets.review);
let onSheet = 0;
revSheet.eachRow((row, n) => {
  if (n >= 4 && row.getCell(1).value instanceof Date) onSheet++;
});
check(onSheet === 1, 'the carried row is ON the Review sheet, not just in the data',
  `${onSheet} dated row(s)`);

const fxSheet = wb.getWorksheet(rules.cards.amex.sheets.fx);
let fxCount = 0;
fxSheet?.eachRow((row, n) => { if (n >= 4 && row.getCell(1).value instanceof Date) fxCount++; });
check(fxCount === 2, 'the USD sheet lists both months', `${fxCount} row(s)`);

// ────────────────────────────────────────────────────── totals must hold ──
console.log('\n--- C. the money still ties out');
const total = after.rows.reduce((s, t) => s + t.amount, 0);
const expected = [...july, ...august].reduce((s, t) => s + t.amount, 0);
check(Math.abs(total - expected) < 0.005, 'merged total matches the sum of both months',
  `${total.toFixed(2)} vs ${expected.toFixed(2)}`);

console.log(failures ? `\n${failures} check(s) failed.` : '\nAll checks passed.');
process.exit(failures ? 1 : 0);
