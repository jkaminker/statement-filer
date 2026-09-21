// Builds the quarterly transactions workbook with ExcelJS, and reads a reviewed
// one back. Layout matches the files already in the audit folder exactly, so a
// new quarter sits beside the old ones without looking different.

const MONEY = '"$"#,##0.00;[Red]("$"#,##0.00)';
const DATE_FMT = 'mmm dd, yyyy';
const BLUE = 'FF1F4E78';

function styleHeader(row, style) {
  row.eachCell((cell) => {
    if (style === 'blueFill') {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLUE } };
    } else {
      cell.font = { bold: true, color: { argb: BLUE } };
    }
    cell.alignment = { horizontal: 'left' };
  });
}

const NOTE_FONT = { name: 'Calibri', size: 10, italic: true, color: { argb: 'FF595959' } };
const BODY_FONT = { name: 'Calibri', size: 11 };

/**
 * @param {object} opts
 *   card, quarter, transactions[{date,desc,amount,category,spendCategory,fx}],
 *   review[], flags[], statements[{label, controlTotal}], rules, notes[]
 */
export async function buildWorkbook(ExcelJS, opts) {
  const { card, quarter, transactions, review, flags, statements, rules, notes = [],
          outside = [], bounds = null } = opts;
  const cfg = rules.cards[card];
  const outsideSheet = (cfg.sheets || {}).outside || 'Outside This Quarter';
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Statement Filer';
  wb.created = new Date();

  const sorted = transactions.slice().sort(
    (a, b) => a.date.localeCompare(b.date) || a.desc.localeCompare(b.desc)
  );

  // Sheets appear in creation order and the summary should be the one that opens,
  // so claim it first and fill it in once the data sheet's row count is known.
  const sum = wb.addWorksheet(cfg.sheets.summary);

  // ------------------------------------------------------------------ Data
  const data = wb.addWorksheet(cfg.sheets.data);
  data.addRow(cfg.dataHeaders);
  styleHeader(data.getRow(1), cfg.headerStyle);

  const isAmex = card === 'amex';
  for (const t of sorted) {
    const row = isAmex
      ? data.addRow([new Date(t.date + 'T00:00:00'), t.desc, t.amount, t.category])
      : data.addRow([new Date(t.date + 'T00:00:00'), t.desc, t.category, t.amount]);
    row.font = BODY_FONT;
    row.getCell(1).numFmt = DATE_FMT;
    row.getCell(isAmex ? 3 : 4).numFmt = MONEY;
  }
  data.views = [{ state: 'frozen', ySplit: 1 }];
  data.autoFilter = { from: 'A1', to: { row: sorted.length + 1, column: 4 } };
  const widths = isAmex ? [13, 45, 13, 22] : [16, 42, 22, 13];
  widths.forEach((w, i) => { data.getColumn(i + 1).width = w; });

  const nRows = sorted.length + 1;
  const catCol = isAmex ? 'D' : 'C';
  const amtCol = isAmex ? 'C' : 'D';
  const dataRef = `'${cfg.sheets.data}'`;

  // --------------------------------------------------------------- Summary
  sum.getCell('A1').value = `${cfg.label} ${quarter} Expense Summary`;
  sum.getCell('A1').font = { bold: true, size: 12 };

  // the out-of-GTA category is named per card: "Travel" on Rogers, "Business
  // Travel" on Amex and CIBC, matching what each folder already uses
  const travelCategory = cfg.gtaToCategory || (rules.gtaRule || {}).toCategory || '';
  const gtaNote = rules.gtaRule && rules.gtaRule.enabled
    ? String(rules.gtaRule.note || '').replace('{travelCategory}', travelCategory) + ' '
    : '';
  const periodLine =
    `Built from ${statements.map((s) => s.label).join(' & ')} statement${statements.length > 1 ? 's' : ''}. `
    + gtaNote
    + notes.join(' ');
  sum.getCell('A2').value = periodLine;
  sum.getCell('A2').font = NOTE_FONT;

  const hdr = sum.getRow(3);
  hdr.values = cfg.summaryHeaders;
  styleHeader(hdr, cfg.headerStyle);

  const cats = [...new Set(sorted.map((t) => t.category))].sort();
  const reviewCat = rules.reviewCategory || 'Review';
  const ordered = [...cats.filter((c) => c !== reviewCat), ...cats.filter((c) => c === reviewCat)];

  let r = 4;
  for (const c of ordered) {
    sum.getCell(`A${r}`).value = c;
    sum.getCell(`A${r}`).font = BODY_FONT;
    sum.getCell(`B${r}`).value = {
      formula: `SUMIF(${dataRef}!$${catCol}$2:$${catCol}$${nRows},$A${r},${dataRef}!$${amtCol}$2:$${amtCol}$${nRows})`,
    };
    sum.getCell(`B${r}`).numFmt = MONEY;
    sum.getCell(`B${r}`).font = BODY_FONT;
    r++;
  }
  sum.getCell(`A${r}`).value = 'Grand Total';
  sum.getCell(`A${r}`).font = { bold: true };
  sum.getCell(`B${r}`).value = { formula: `SUM(B4:B${r - 1})` };
  sum.getCell(`B${r}`).numFmt = MONEY;
  sum.getCell(`B${r}`).font = { bold: true };
  const totalRow = r;

  // ------------------------------------------------- reconciliation block
  r += 2;
  sum.getCell(`A${r}`).value = 'Reconciliation to statements';
  sum.getCell(`A${r}`).font = { bold: true };
  r++;
  const firstStmtRow = r;
  for (const s of statements) {
    sum.getCell(`A${r}`).value = cfg.controlTotalLabel.replace('{date}', s.label);
    sum.getCell(`A${r}`).font = BODY_FONT;
    sum.getCell(`B${r}`).value = s.controlTotal;
    sum.getCell(`B${r}`).numFmt = MONEY;
    sum.getCell(`B${r}`).font = BODY_FONT;
    r++;
  }
  sum.getCell(`A${r}`).value = 'Statements combined';
  sum.getCell(`A${r}`).font = { bold: true };
  sum.getCell(`B${r}`).value = { formula: `SUM(B${firstStmtRow}:B${r - 1})` };
  sum.getCell(`B${r}`).numFmt = MONEY;
  sum.getCell(`B${r}`).font = { bold: true };
  const stmtRow = r;
  r++;

  // A statement period is not a quarter, so the Grand Total above covers only
  // the quarter while the statements cover their own periods. The two tie only
  // once the out-of-quarter rows are added back, and showing that sum here is
  // what keeps the variance check meaningful: it still proves every line on
  // every statement was parsed, even though the summary is scoped to a quarter.
  const outsideTotal = round2(outside.reduce((s, t) => s + t.amount, 0));
  sum.getCell(`A${r}`).value = `Grand Total (${quarter} only)`;
  sum.getCell(`A${r}`).font = BODY_FONT;
  sum.getCell(`B${r}`).value = { formula: `B${totalRow}` };
  sum.getCell(`B${r}`).numFmt = MONEY;
  sum.getCell(`B${r}`).font = BODY_FONT;
  const inQuarterRow = r;
  r++;
  sum.getCell(`A${r}`).value = `Outside ${quarter} (see '${outsideSheet}')`;
  sum.getCell(`A${r}`).font = BODY_FONT;
  sum.getCell(`B${r}`).value = outsideTotal;
  sum.getCell(`B${r}`).numFmt = MONEY;
  sum.getCell(`B${r}`).font = BODY_FONT;
  const outsideRow = r;
  r++;
  sum.getCell(`A${r}`).value = 'Quarter + outside';
  sum.getCell(`A${r}`).font = { bold: true };
  sum.getCell(`B${r}`).value = { formula: `B${inQuarterRow}+B${outsideRow}` };
  sum.getCell(`B${r}`).numFmt = MONEY;
  sum.getCell(`B${r}`).font = { bold: true };
  const accountedRow = r;
  r++;
  sum.getCell(`A${r}`).value = 'Variance vs. statements';
  sum.getCell(`A${r}`).font = { bold: true };
  sum.getCell(`B${r}`).value = { formula: `B${accountedRow}-B${stmtRow}` };
  sum.getCell(`B${r}`).numFmt = MONEY;
  sum.getCell(`B${r}`).font = { bold: true };

  sum.getColumn(1).width = 34;
  sum.getColumn(2).width = 18;

  // ---------------------------------------------------------------- Review
  const rev = wb.addWorksheet(cfg.sheets.review);
  rev.getCell('A1').value = 'Items to review & assign a category:';
  rev.getCell('A1').font = { bold: true };
  rev.getCell('A2').value =
    `These are coded '${reviewCat}' on the ${cfg.sheets.data} sheet and are excluded from every `
    + 'other category in the summary. Put your answer in the COMMENTS column, save the file, '
    + 'then load it back into the app.';
  rev.getCell('A2').font = NOTE_FONT;

  const revHdr = rev.getRow(3);
  revHdr.values = ['TRANSACTION DATE', 'TRANSACTION DETAILS', 'SUGGESTED CATEGORY',
                   'AMOUNT', 'NOTES', 'COMMENTS'];
  styleHeader(revHdr, cfg.headerStyle);

  let rr = 4;
  for (const item of review.slice().sort((a, b) => a.date.localeCompare(b.date))) {
    const row = rev.getRow(rr);
    // Rogers keeps the town in its own field, so it isn't in the description you
    // see here. Lead the note with it — without it you'd be choosing a category
    // for "TIM HORTONS #3025" with no way to know it was in Bracebridge.
    const note = item.place
      ? `${item.place}. ${item.note || ''}`.trim()
      : (item.note || '');
    row.values = [new Date(item.date + 'T00:00:00'), item.desc, item.suggested || '',
                  item.amount, note, ''];
    row.font = BODY_FONT;
    row.getCell(1).numFmt = DATE_FMT;
    row.getCell(4).numFmt = MONEY;
    row.getCell(5).alignment = { wrapText: true, vertical: 'top' };
    rr++;
  }
  if (review.length) {
    rev.getCell(`B${rr}`).value = 'Total in Review';
    rev.getCell(`B${rr}`).font = { bold: true };
    rev.getCell(`D${rr}`).value = { formula: `SUM(D4:D${rr - 1})` };
    rev.getCell(`D${rr}`).numFmt = MONEY;
    rev.getCell(`D${rr}`).font = { bold: true };
  } else {
    // no SUM here: an empty list would make the range run backwards and yield #VALUE!
    rev.getCell(`A${rr}`).value =
      'Nothing needed a decision this time — every merchant matched a rule you already have.';
    rev.getCell(`A${rr}`).font = NOTE_FONT;
  }

  if (flags.length) {
    rr += 3;
    rev.getCell(`A${rr}`).value = 'Categorized, but flagged for a second look:';
    rev.getCell(`A${rr}`).font = { bold: true };
    rr++;
    const fh = rev.getRow(rr);
    fh.values = ['MERCHANT', 'ASSIGNED CATEGORY', '', 'AMOUNT', 'NOTES', 'COMMENTS'];
    styleHeader(fh, cfg.headerStyle);
    rr++;
    for (const f of flags) {
      const row = rev.getRow(rr);
      row.values = [f.desc, f.category, '', f.amount, f.note, ''];
      row.font = BODY_FONT;
      row.getCell(4).numFmt = MONEY;
      row.getCell(5).alignment = { wrapText: true, vertical: 'top' };
      rr++;
    }
  }

  [16, 42, 22, 13, 70, 24].forEach((w, i) => { rev.getColumn(i + 1).width = w; });

  // ------------------------------------------------- outside this quarter
  // Charges that arrived on these statements but belong to a different quarter.
  // They are kept — nothing is thrown away — but they are off the Data sheet,
  // so no category total includes them.
  if (outside.length) {
    const out = wb.addWorksheet(outsideSheet);
    out.getCell('A1').value = `Transactions on these statements that fall outside ${quarter}`;
    out.getCell('A1').font = { bold: true, size: 12 };
    out.getCell('A2').value =
      (bounds ? `${quarter} runs ${bounds.from} to ${bounds.to}. ` : '')
      + 'A statement period is not a quarter - the first statement of a quarter carries '
      + 'charges from the end of the previous one. These rows are excluded from the '
      + `${cfg.sheets.summary} totals and belong in the quarter named in the last column.`;
    out.getCell('A2').font = NOTE_FONT;
    out.getCell('A2').alignment = { wrapText: true, vertical: 'top' };

    const oh = out.getRow(4);
    oh.values = ['Date', 'Description', 'Amount', 'Category', 'Belongs to'];
    styleHeader(oh, cfg.headerStyle);

    let orow = 5;
    for (const t of outside) {
      const row = out.getRow(orow);
      row.values = [new Date(t.date + 'T00:00:00'), t.desc, t.amount,
                    t.category || '', quarterLabelOf(t.date)];
      row.font = BODY_FONT;
      row.getCell(1).numFmt = DATE_FMT;
      row.getCell(3).numFmt = MONEY;
      orow++;
    }
    out.getCell(`B${orow}`).value = 'Total outside this quarter';
    out.getCell(`B${orow}`).font = { bold: true };
    out.getCell(`C${orow}`).value = { formula: `SUM(C5:C${orow - 1})` };
    out.getCell(`C${orow}`).numFmt = MONEY;
    out.getCell(`C${orow}`).font = { bold: true };

    [13, 45, 13, 22, 14].forEach((w, i) => { out.getColumn(i + 1).width = w; });
    out.getRow(2).height = 30;
  }

  // ----------------------------------------------- foreign-currency detail
  const fxRows = sorted.filter((t) => t.fx);
  if (cfg.sheets.fx && fxRows.length) {
    const fx = wb.addWorksheet(cfg.sheets.fx);
    fx.getCell('A1').value =
      'Foreign-currency charges - reference only. Every row below is also included in '
      + `${cfg.sheets.data}; do not add these to the summary.`;
    fx.getCell('A1').font = NOTE_FONT;
    const h = fx.getRow(3);
    h.values = ['Date', 'Description', 'Amount (CAD)', 'Foreign Spend Amount',
                'Exchange Rate', 'Category'];
    styleHeader(h, cfg.headerStyle);
    let fr = 4;
    for (const t of fxRows) {
      const row = fx.getRow(fr);
      row.values = [new Date(t.date + 'T00:00:00'), t.desc, t.amount,
                    `${t.fx.amount.toFixed(2)} ${t.fx.currency}`, t.fx.rate, t.category];
      row.font = BODY_FONT;
      row.getCell(1).numFmt = DATE_FMT;
      row.getCell(3).numFmt = MONEY;
      fr++;
    }
    [13, 42, 14, 28, 14, 22].forEach((w, i) => { fx.getColumn(i + 1).width = w; });
  }

  return wb;
}

/**
 * Read a reviewed workbook: pull the COMMENTS column from both blocks of the
 * Review sheet, and return the decisions plus the existing Data rows.
 */
export async function readReviewedWorkbook(ExcelJS, arrayBuffer, card, rules) {
  const cfg = rules.cards[card];
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);

  const rev = wb.getWorksheet(cfg.sheets.review);
  const data = wb.getWorksheet(cfg.sheets.data);
  if (!rev || !data) throw new Error('That workbook does not have the expected sheets.');

  const rowDecisions = [];    // keyed by date+desc+amount
  const merchantDecisions = []; // keyed by merchant text (the flagged block)

  rev.eachRow((row, n) => {
    if (n < 4) return;
    const a = row.getCell(1).value;
    const desc = row.getCell(2).value;
    const amount = row.getCell(4).value;
    const comment = row.getCell(6).value;
    if (!comment || !String(comment).trim()) return;
    const category = String(comment).trim();
    if (a instanceof Date) {
      rowDecisions.push({
        date: a.toISOString().slice(0, 10),
        desc: String(desc || '').trim(),
        amount: Math.round(Number(amount) * 100) / 100,
        category,
      });
    } else if (typeof a === 'string' && a.trim() && a.trim() !== 'MERCHANT') {
      merchantDecisions.push({ desc: a.trim(), category });
    }
  });

  const isAmex = card === 'amex';
  const rows = [];
  data.eachRow((row, n) => {
    if (n === 1) return;
    const d = row.getCell(1).value;
    if (!(d instanceof Date)) return;
    rows.push({
      date: d.toISOString().slice(0, 10),
      desc: String(row.getCell(2).value || ''),
      category: String(row.getCell(isAmex ? 4 : 3).value || ''),
      amount: Number(row.getCell(isAmex ? 3 : 4).value),
    });
  });

  return { rowDecisions, merchantDecisions, rows };
}

/**
 * Read a workbook this app filed in an earlier run, so a later drop can merge
 * into it instead of replacing it.
 *
 * Four things come back. The Data rows carry categories you have already settled
 * — they are taken as final and never re-categorized. The reconciliation block
 * tells us which statements the quarter already covers, and that list is what
 * makes re-dropping a statement harmless: a statement already named there is
 * skipped whole rather than having its rows matched one by one. The FX sheet is
 * read back and reattached to its rows, and the Review sheet's notes are
 * recovered, because neither survives a round trip through the Data sheet alone:
 * a row that is still awaiting your decision would otherwise come back stripped
 * of the note that tells you what you are deciding about.
 */
export async function readFiledWorkbook(ExcelJS, arrayBuffer, card, rules) {
  const cfg = rules.cards[card];
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);

  const data = wb.getWorksheet(cfg.sheets.data);
  const sum = wb.getWorksheet(cfg.sheets.summary);
  if (!data) throw new Error(`That workbook has no "${cfg.sheets.data}" sheet.`);

  const isAmex = card === 'amex';

  // ------------------------------------------- foreign-currency detail, first
  // Column D holds "123.45 USD" and column E the rate. Keyed by date+desc+amount
  // so it can be reattached to the Data row it belongs to.
  const fxByKey = new Map();
  const fxSheet = cfg.sheets.fx ? wb.getWorksheet(cfg.sheets.fx) : null;
  if (fxSheet) {
    fxSheet.eachRow((row, n) => {
      if (n < 4) return;
      const d = row.getCell(1).value;
      if (!(d instanceof Date)) return;
      const amount = cellNumber(row.getCell(3).value);
      const spend = String(cellText(row.getCell(4).value) || '').trim();
      const rate = cellNumber(row.getCell(5).value);
      const m = spend.match(/^(-?[\d,]+(?:\.\d+)?)\s*([A-Za-z]{3})$/);
      if (!m || !Number.isFinite(amount)) return;
      fxByKey.set(
        rowKey(isoOf(d), cellText(row.getCell(2).value), amount),
        { amount: Number(m[1].replace(/,/g, '')), currency: m[2].toUpperCase(), rate }
      );
    });
  }

  // --------------------------------------------- the Review sheet's own notes
  // Only the top block (dated rows). The flagged block below it is keyed by
  // merchant, not by row, and is rebuilt from scratch on every run anyway.
  const reviewByKey = new Map();
  const revSheet = wb.getWorksheet(cfg.sheets.review);
  if (revSheet) {
    revSheet.eachRow((row, n) => {
      if (n < 4) return;
      const d = row.getCell(1).value;
      if (!(d instanceof Date)) return;
      const amount = cellNumber(row.getCell(4).value);
      if (!Number.isFinite(amount)) return;
      reviewByKey.set(rowKey(isoOf(d), cellText(row.getCell(2).value), amount), {
        suggested: String(cellText(row.getCell(3).value) || '').trim(),
        note: String(cellText(row.getCell(5).value) || '').trim(),
      });
    });
  }

  // ------------------------------------------------ outside this quarter
  // Rows a previous run set aside because they belong to another quarter. They
  // have to come back or the next drop would rebuild the workbook without them
  // and the reconciliation would stop tying to the statements.
  const outside = [];
  const outSheet = wb.getWorksheet((cfg.sheets || {}).outside || 'Outside This Quarter');
  if (outSheet) {
    outSheet.eachRow((row, n) => {
      if (n < 5) return;
      const d = row.getCell(1).value;
      if (!(d instanceof Date)) return;
      const amount = cellNumber(row.getCell(3).value);
      if (!Number.isFinite(amount)) return;
      outside.push({
        date: isoOf(d),
        desc: String(cellText(row.getCell(2).value) || ''),
        amount: Math.round(amount * 100) / 100,
        category: String(cellText(row.getCell(4).value) || ''),
        filed: true,
        outside: true,
      });
    });
  }

  const rows = [];
  data.eachRow((row, n) => {
    if (n === 1) return;
    const d = row.getCell(1).value;
    if (!(d instanceof Date)) return;
    const amount = cellNumber(row.getCell(isAmex ? 3 : 4).value);
    if (!Number.isFinite(amount)) return;
    const date = isoOf(d);
    const desc = String(cellText(row.getCell(2).value) || '');
    const rounded = Math.round(amount * 100) / 100;
    const key = rowKey(date, desc, rounded);
    const t = {
      date,
      desc,
      category: String(cellText(row.getCell(isAmex ? 4 : 3).value) || ''),
      amount: rounded,
      filed: true,
    };
    const fx = fxByKey.get(key);
    if (fx) t.fx = fx;
    const rev = reviewByKey.get(key);
    if (rev) {
      if (rev.suggested) t.suggested = rev.suggested;
      if (rev.note) t.note = rev.note;
    }
    rows.push(t);
  });

  // pull the per-statement control totals back out of the reconciliation block
  // by running the card's label template backwards
  const statements = [];
  if (sum) {
    const tpl = String(cfg.controlTotalLabel || '{date} statement');
    const re = new RegExp('^' + tpl.split('{date}').map(escapeRe).join('(.+)') + '$');
    sum.eachRow((row) => {
      const label = String(row.getCell(1).value || '').trim();
      const m = label.match(re);
      if (!m) return;
      const v = row.getCell(2).value;
      const controlTotal = Number(v && typeof v === 'object' ? v.result : v);
      if (Number.isFinite(controlTotal)) statements.push({ label: m[1].trim(), controlTotal });
    });
  }
  return { rows, statements, outside };
}

function escapeRe(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

/** Quarter label for an ISO date — kept local so this module stands alone. */
function quarterLabelOf(iso) {
  const [y, m] = String(iso).split('-').map(Number);
  return `Q${Math.floor((m - 1) / 3) + 1} ${y}`;
}

// ─────────────────────────────────────────────────────────── cell helpers ──
// A cell that looks like a plain number or string can arrive as a formula
// result or as rich text, depending on how Excel last saved the file. These
// three flatten all of that down to what the row actually says.

function cellText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((r) => r.text).join('');
    if ('result' in v) return String(v.result ?? '');
    if ('text' in v) return String(v.text ?? '');
  }
  return String(v);
}

function cellNumber(v) {
  if (typeof v === 'number') return v;
  if (v && typeof v === 'object' && 'result' in v) return Number(v.result);
  const n = Number(String(cellText(v)).replace(/[$,\s]/g, ''));
  return Number.isFinite(n) ? n : NaN;
}

function isoOf(d) {
  // matches how buildWorkbook wrote the date, so a value survives the round trip
  return d.toISOString().slice(0, 10);
}

/**
 * The identity of a transaction across sheets and across runs: the same date,
 * merchant and amount. Whitespace is normalized because Excel and the parsers
 * disagree about non-breaking spaces.
 */
export function rowKey(date, desc, amount) {
  const d = String(desc || '').replace(/ /g, ' ').replace(/\s+/g, ' ').trim().toUpperCase();
  return `${date}|${d}|${Math.round(Number(amount) * 100) / 100}`;
}
