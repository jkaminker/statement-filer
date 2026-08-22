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
  const { card, quarter, transactions, review, flags, statements, rules, notes = [] } = opts;
  const cfg = rules.cards[card];
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
  sum.getCell(`A${r}`).value = 'Variance vs. Grand Total';
  sum.getCell(`A${r}`).font = { bold: true };
  sum.getCell(`B${r}`).value = { formula: `B${totalRow}-B${stmtRow}` };
  sum.getCell(`B${r}`).numFmt = MONEY;
  sum.getCell(`B${r}`).font = { bold: true };

  sum.getColumn(1).width = 30;
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
    row.values = [new Date(item.date + 'T00:00:00'), item.desc, item.suggested || '',
                  item.amount, item.note || '', ''];
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
