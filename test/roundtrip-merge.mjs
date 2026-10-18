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

// ───────────────────────────────── the quarter is a quarter, not a statement ──
// The July statement runs from roughly 18 June to 17 July, so it always carries
// June charges. Those belong to Q2 and must not reach a single Q3 category.
console.log('\n--- D. June charges on the July statement stay out of Q3');
const { quarterBounds, inQuarter } = await import(path.join(ROOT, 'js/parsers/base.js'));

const b = quarterBounds('Q3 2026');
check(b.from === '2026-07-01' && b.to === '2026-09-30', 'Q3 2026 runs Jul 1 - Sep 30',
  `${b.from} .. ${b.to}`);
const b4 = quarterBounds('Q4 2026');
check(b4.from === '2026-10-01' && b4.to === '2026-12-31', 'Q4 2026 runs Oct 1 - Dec 31',
  `${b4.from} .. ${b4.to}`);
const b1 = quarterBounds('Q1 2024');
check(b1.to === '2024-03-31', 'a leap year still ends Q1 on Mar 31', b1.to);
check(!inQuarter('2026-06-30', 'Q3 2026'), 'Jun 30 is outside Q3');
check(inQuarter('2026-07-01', 'Q3 2026'), 'Jul 1 is inside Q3');
check(inQuarter('2026-09-30', 'Q3 2026'), 'Sep 30 is inside Q3');
check(!inQuarter('2026-10-01', 'Q3 2026'), 'Oct 1 is outside Q3');

// a realistic July statement: a fortnight of June, then July
const julyStatement = [
  { date: '2026-06-21', desc: 'ESSO CIRCLE K', amount: 74.10, category: 'Fuel' },
  { date: '2026-06-28', desc: 'LCBO #217', amount: 43.85, category: 'Meals' },
  { date: '2026-07-02', desc: 'PRESTO FARE', amount: 3.35, category: 'Travel' },
  { date: '2026-07-14', desc: 'STAPLES #118', amount: 128.99, category: 'Office Supplies' },
];
const inQ = julyStatement.filter((t) => inQuarter(t.date, 'Q3 2026'));
const outQ = julyStatement.filter((t) => !inQuarter(t.date, 'Q3 2026'));
check(inQ.length === 2 && outQ.length === 2, 'the statement splits 2 in / 2 out',
  `${inQ.length} in, ${outQ.length} out`);

const stmtTotal = round2(julyStatement.reduce((s, t) => s + t.amount, 0));
const q3Bytes2 = await bytesOf({
  card: 'amex', quarter: 'Q3 2026', transactions: inQ, review: [], flags: [],
  statements: [{ label: 'Jul 17 2026', controlTotal: stmtTotal }],
  rules, notes: [], outside: outQ, bounds: b,
});

const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.load(q3Bytes2);

// the summary must not know about June at all
const sum2 = wb2.getWorksheet(rules.cards.amex.sheets.summary);
const catCells = [];
sum2.eachRow((row, n) => { if (n >= 4) catCells.push(String(row.getCell(1).value || '')); });
check(!catCells.includes('Fuel'), 'Fuel (a June-only category) is absent from the summary',
  catCells.filter(Boolean).slice(0, 8).join(', '));

const data2 = wb2.getWorksheet(rules.cards.amex.sheets.data);
let dataDates = [];
data2.eachRow((row, n) => {
  if (n === 1) return;
  const d = row.getCell(1).value;
  if (d instanceof Date) dataDates.push(d.toISOString().slice(0, 10));
});
check(dataDates.every((d) => d >= '2026-07-01'), 'no June row reaches the Data sheet',
  dataDates.join(', '));
check(dataDates.length === 2, 'the Data sheet holds exactly the two July rows',
  `${dataDates.length}`);

const outSheet = wb2.getWorksheet('Outside This Quarter');
check(!!outSheet, 'the outside-quarter sheet exists');
let outRows = 0;
let outSum = 0;
outSheet?.eachRow((row, n) => {
  if (n < 5) return;
  const d = row.getCell(1).value;
  if (!(d instanceof Date)) return;
  outRows++;
  outSum += Number(row.getCell(3).value) || 0;
});
check(outRows === 2, 'both June rows are listed there', `${outRows}`);
check(Math.abs(round2(outSum) - 117.95) < 0.005, 'and they total the June amount',
  round2(outSum).toFixed(2));

// the reconciliation identity: quarter + outside === statements
const inTotal = round2(inQ.reduce((s, t) => s + t.amount, 0));
const outTotal = round2(outQ.reduce((s, t) => s + t.amount, 0));
check(Math.abs(inTotal + outTotal - stmtTotal) < 0.005,
  'quarter + outside ties to the statement control total',
  `${inTotal} + ${outTotal} vs ${stmtTotal}`);

// and it survives a read back, so the next drop doesn't lose the June rows
const back2 = await readFiledWorkbook(ExcelJS, q3Bytes2, 'amex', rules);
check(back2.rows.length === 2, 'reading back gives 2 in-quarter rows', `${back2.rows.length}`);
check((back2.outside || []).length === 2, 'and 2 outside rows',
  `${(back2.outside || []).length}`);

// ─────────────────────────── identical charges are not the same transaction ──
// Two Presto fares on one morning, two coffees at one shop: same date, same
// merchant, same amount, two real charges. Collapsing them destroys money and
// breaks the reconciliation, so the out-of-quarter list must keep both.
console.log('\n--- D2. duplicate-looking charges both survive');
const twins = [
  { date: '2026-06-21', desc: 'PRESTO FARE', amount: 55.37, category: 'Travel' },
  { date: '2026-06-21', desc: 'PRESTO FARE', amount: 55.37, category: 'Travel' },
  { date: '2026-07-06', desc: 'STAPLES #118', amount: 40.00, category: 'Office Supplies' },
];
const twinsIn = twins.filter((t) => inQuarter(t.date, 'Q3 2026'));
const twinsOut = twins.filter((t) => !inQuarter(t.date, 'Q3 2026'));
const twinsControl = round2(twins.reduce((s, t) => s + t.amount, 0));

const twinBytes = await bytesOf({
  card: 'amex', quarter: 'Q3 2026', transactions: twinsIn, review: [], flags: [],
  statements: [{ label: 'Jul 17 2026', controlTotal: twinsControl }],
  rules, notes: [], outside: twinsOut, bounds: b,
});
const twinBack = await readFiledWorkbook(ExcelJS, twinBytes, 'amex', rules);
check(twinBack.outside.length === 2, 'both identical fares are on the outside sheet',
  `${twinBack.outside.length} row(s)`);
const twinOutTotal = round2(twinBack.outside.reduce((s, t) => s + t.amount, 0));
check(Math.abs(twinOutTotal - 110.74) < 0.005, 'and they total both fares, not one',
  `$${twinOutTotal.toFixed(2)} (one fare alone would be $55.37)`);
const twinAll = round2(twinBack.rows.reduce((s, t) => s + t.amount, 0)) + twinOutTotal;
check(Math.abs(twinAll - twinsControl) < 0.005,
  'quarter + outside still ties to the statement',
  `${twinAll.toFixed(2)} vs ${twinsControl.toFixed(2)}`);

// ──────────────────────────────────────────── statement naming convention ──
console.log('\n--- E. statements are filed under your naming convention');
const { statementFileName } = await import(path.join(ROOT, 'js/pipeline.js'));
const amexName = statementFileName(rules, 'amex', { statementDate: '2026-07-17' });
check(amexName === 'Amex Jul 17 2026 Statement.pdf', 'Amex', amexName);
const cibcName = statementFileName(rules, 'cibc', { statementDate: '2026-06-13' });
check(cibcName === 'CIBC Statement - Jun 13 2026.pdf', 'CIBC', cibcName);
const singleDigit = statementFileName(rules, 'amex', { statementDate: '2026-08-05' });
check(singleDigit === 'Amex Aug 5 2026 Statement.pdf', 'a single-digit day is not padded',
  singleDigit);

// ─────────────────────────── the fiscal-year folder follows the quarter ──
// An October statement closes out Q3 but belongs to the NEXT fiscal year.
// Dating the audit folder from the statement sent Q3 2026 lookups into
// "Annual Audit 2027 Sep", which does not exist — so the app reported the
// quarter as unfiled, greyed out "Add to this quarter", and offered to
// rebuild it from the single statement in hand.
console.log('\n--- F. the fiscal-year folder follows the quarter, not the statement');
const { fiscalYearOfQuarter, fiscalYearOf, monthNum } = await import(
  path.join(ROOT, 'js/parsers/base.js')
);
const FYE = rules.fiscalYearEndMonth;
check(FYE === 9, 'fiscal year ends September', String(FYE));
check(fiscalYearOfQuarter('Q3 2026', FYE) === 2026,
  'Q3 2026 files under FY2026', String(fiscalYearOfQuarter('Q3 2026', FYE)));
check(fiscalYearOfQuarter('Q4 2026', FYE) === 2027,
  'Q4 2026 files under FY2027', String(fiscalYearOfQuarter('Q4 2026', FYE)));
check(fiscalYearOfQuarter('Q1 2026', FYE) === 2026,
  'Q1 2026 files under FY2026', String(fiscalYearOfQuarter('Q1 2026', FYE)));
check(fiscalYearOf('2026-10-13', FYE) === 2027,
  'an Oct 13 2026 statement is itself FY2027 — the old, wrong basis');
check(fiscalYearOfQuarter('Q3 2026', FYE) !== fiscalYearOf('2026-10-13', FYE),
  'which is exactly why the two must not be the same number');
check(fiscalYearOfQuarter('nonsense', FYE) === null,
  'an unparseable quarter yields null so the caller can fall back');

// ──────────────────────────────── statement months, abbreviated or full ──
// CIBC writes "September 14 to October 13, 2026". The abbreviation map alone
// returned undefined, which reached Drive as "CIBC Statement - undefined 13
// 2026.pdf" and as an undefined month inside the statement date.
console.log('\n--- G. month names parse in both forms');
check(monthNum('Oct') === 10, 'abbreviated "Oct"', String(monthNum('Oct')));
check(monthNum('October') === 10, 'full "October"', String(monthNum('October')));
check(monthNum('September') === 9, 'full "September"', String(monthNum('September')));
check(monthNum('MARCH') === 3, 'case-insensitive full name', String(monthNum('MARCH')));
check(monthNum('Marketing') === undefined,
  'a merchant word is NOT read as a month', String(monthNum('Marketing')));
check(monthNum('Mayfair') === undefined,
  'nor is "Mayfair" read as May', String(monthNum('Mayfair')));
check(monthNum(undefined) === undefined, 'undefined in, undefined out');

// ─────────────────────────────────── what actually reaches the ledger ──
// planEntries decides what gets written into Wave. Everything the preview
// shows comes from it, so these checks are the difference between a correct
// quarter and a wrong one posted confidently.
console.log('\n--- H. the Wave posting plan');
const { planEntries } = await import(path.join(ROOT, 'js/wave.js'));

const waveRun = {
  card: 'amex',
  cardLabel: 'Amex',
  quarter: 'Q3 2026',
  summary: {
    'Business Travel': 21702.77,
    'Travel - Domestic': 14454.65,
    'Travel': 2456.44,
    'Professional Fees': 9839.12,
    'Meals': 3588.65,
    'Rent': 1826.09,
    'Personal': 11286.00,
    'Personal - Home Insurance': 1096.71,
  },
};
const plan = planEntries(waveRun, rules);
const byAccount = {};
for (const e of plan.entries) byAccount[e.account] = round2((byAccount[e.account] || 0) + e.amount);

check(plan.unmapped.length === 0, 'every business category maps to a Wave account',
  plan.unmapped.map((u) => u.category).join(', ') || 'none unmapped');
check(plan.total === 53867.72, 'the posted total is business-only', `$${plan.total}`);
check(round2(plan.excluded.reduce((s, x) => s + x.amount, 0)) === 12382.71,
  'personal is held out, not netted', `$${round2(plan.excluded.reduce((s, x) => s + x.amount, 0))}`);
check(!plan.entries.some((e) => /Personal/.test(e.category)),
  'no personal category reaches the ledger');
check(byAccount['Business Travel'] === 24159.21,
  'plain Travel folds into Business Travel', `$${byAccount['Business Travel']}`);
check(byAccount['Meals and Entertainment'] === 3588.65,
  'Meals maps to the Wave account name', `$${byAccount['Meals and Entertainment']}`);
check(plan.entries.every((e) => e.date === '2026-09-30'),
  'every line is dated the quarter end');
check(plan.entries.every((e) => e.direction === 'WITHDRAWAL' && e.balance === 'DEBIT'),
  'expenses are withdrawals debiting the expense account');

// Re-planning the same run must produce the same ids, or a second post would
// duplicate the quarter instead of updating it.
const again = planEntries(waveRun, rules);
check(JSON.stringify(plan.entries.map((e) => e.externalId))
   === JSON.stringify(again.entries.map((e) => e.externalId)),
  'ids are stable across runs, so re-posting cannot duplicate');
check(new Set(plan.entries.map((e) => e.externalId)).size === plan.entries.length,
  'and unique within the run');
check(plan.entries.some((e) => e.externalId === 'sf-amex-q3-2026-business-travel'),
  'ids are readable', plan.entries[0].externalId);

console.log('\n--- I. an unmapped category blocks the post');
const oddRun = { ...waveRun, summary: { 'Professional Fees': 100, 'Hot Air Ballooning': 250.00 } };
const oddPlan = planEntries(oddRun, rules);
check(oddPlan.unmapped.length === 1 && oddPlan.unmapped[0].category === 'Hot Air Ballooning',
  'the unknown category is reported, not silently dropped');
check(!oddPlan.entries.some((e) => e.category === 'Hot Air Ballooning'),
  'and it is not posted to some fallback account');
check(oddPlan.total === 100, 'the rest still totals correctly', `$${oddPlan.total}`);

console.log('\n--- J. a net-negative category reverses direction');
const refundPlan = planEntries({ ...waveRun, summary: { 'Professional Fees': -42.50 } }, rules);
check(refundPlan.entries[0].direction === 'DEPOSIT', 'a net credit is a deposit');
check(refundPlan.entries[0].balance === 'CREDIT', 'and credits the expense account');
check(refundPlan.entries[0].amount === 42.50, 'posted as a positive amount',
  `$${refundPlan.entries[0].amount}`);
check(refundPlan.total === -42.50, 'but counts against the total', `$${refundPlan.total}`);

function round2(n) { return Math.round(n * 100) / 100; }

console.log(failures ? `\n${failures} check(s) failed.` : '\nAll checks passed.');
process.exit(failures ? 1 : 0);
