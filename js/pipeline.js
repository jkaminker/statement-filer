// Orchestrates a run: statements in, categorized workbook + highlighted PDFs out.

import { readPages, quarterOf, fiscalYearOf } from './parsers/base.js';
import { detectCard, parserFor } from './parsers/registry.js';
import { categorize, summarize, applyGtaRule, canonicalCategory } from './rules.js';
import { buildWorkbook, readFiledWorkbook } from './workbook.js';
import { buildCategoryPdfs, buildAbridgedStatements } from './highlight.js';

/**
 * @param {File[]} files          statement PDFs the user dropped
 * @param {object} rules
 * @param {object} libs           {pdfjsLib, ExcelJS, PDFLib}
 * @param {function} onProgress   (message) => void
 */
export async function run(files, rules, libs, onProgress = () => {}, opts = {}) {
  const { pdfjsLib, ExcelJS, PDFLib } = libs;
  const sources = [];

  // ------------------------------------------------------------- 1. parse
  for (const file of files) {
    onProgress(`Reading ${file.name}…`);
    const bytes = new Uint8Array(await file.arrayBuffer());
    // pdf.js consumes (detaches) the buffer it is given, so hand it a copy and
    // keep the original for pdf-lib to use when building the highlighted PDFs
    const { pages } = await readPages(pdfjsLib, bytes.slice().buffer);
    const parser = detectCard(pages);
    if (!parser) {
      throw new Error(
        `${file.name}: I don't recognize this statement format yet. `
        + `Supported so far: ${Object.values(rules.cards).map((c) => c.label).join(', ')}.`
      );
    }
    const parsed = parser.parse(pages);
    if (!parsed.transactions.length) {
      throw new Error(`${file.name}: parsed cleanly but found no transactions.`);
    }
    sources.push({ name: file.name, bytes, pagesMeta: pages, parser, parsed });
    onProgress(
      `${file.name}: ${parser.label}, ${parsed.transactions.length} transactions, `
      + `statement ${parsed.statementLabel}`
    );
  }

  // ------------------------------------------- 2. group by card and quarter
  const cards = [...new Set(sources.map((s) => s.parser.id))];
  if (cards.length > 1) {
    throw new Error(
      'Those statements are from different cards ('
      + cards.map((c) => rules.cards[c].label).join(', ')
      + '). Run one card at a time so each lands in its own folder.'
    );
  }
  const card = cards[0];

  const all = [];
  for (const s of sources) {
    for (const t of s.parsed.transactions) all.push({ ...t, source: s.name });
  }
  all.sort((a, b) => a.date.localeCompare(b.date) || a.desc.localeCompare(b.desc));

  const quarters = [...new Set(all.map((t) => quarterOf(t.date)))];
  const quarter = dominant(all.map((t) => quarterOf(t.date)));
  const fiscalYear = fiscalYearOf(all[all.length - 1].date, rules.fiscalYearEndMonth);

  // ------------------------------------------- 2b. merge with what's filed
  // A quarter is built up over several drops. If this card and quarter already
  // have a workbook, its rows and statements come along so the new statement is
  // ADDED to the quarter rather than replacing it. De-duplication is by
  // statement, not by row: a statement already named in the filed
  // reconciliation block is skipped whole, which makes re-dropping one a no-op.
  let filedRows = [];
  let filedStatements = [];
  let basePdfs = null;
  const skipped = [];
  if (typeof opts.fetchFiled === 'function') {
    try {
      const filed = await opts.fetchFiled(card, quarter, fiscalYear);
      if (filed && filed.workbookBytes) {
        const prior = await readFiledWorkbook(ExcelJS, filed.workbookBytes, card, rules);
        filedRows = prior.rows;
        filedStatements = prior.statements;
        basePdfs = filed.categoryPdfs || null;
        const already = new Set(filedStatements.map((x) => x.label));
        for (let i = sources.length - 1; i >= 0; i--) {
          if (already.has(sources[i].parsed.statementLabel)) {
            skipped.push(sources[i].parsed.statementLabel);
            sources.splice(i, 1);
          }
        }
        for (const label of skipped) {
          onProgress(`${label} is already in the filed workbook — leaving those rows alone.`);
        }
      }
    } catch (e) {
      onProgress(`Could not read what's already filed (${e.message}). Building this quarter fresh.`);
    }
  }

  const fresh = [];
  for (const s of sources) {
    for (const t of s.parsed.transactions) fresh.push({ ...t, source: s.name });
  }

  // --------------------------------------------------------- 3. categorize
  // only the new rows: anything already filed carries the category you settled
  onProgress('Categorizing…');
  const { review, flags } = categorize(fresh, card, rules);
  const merged = [...filedRows, ...fresh].sort(
    (a, b) => a.date.localeCompare(b.date) || a.desc.localeCompare(b.desc)
  );

  // --------------------------------------------------- 4. reconcile & check
  const parsedTotal = round2(merged.reduce((s, t) => s + t.amount, 0));
  const statements = [
    ...filedStatements,
    ...sources.map((s) => ({
      label: s.parsed.statementLabel,
      controlTotal: s.parsed.controlTotal,
    })),
  ];
  const controlTotal = round2(
    statements.reduce((s, x) => s + (x.controlTotal || 0), 0)
  );
  const variance = round2(parsedTotal - controlTotal);

  // --------------------------------------------------------- 5. build files
  onProgress('Building the workbook…');
  const notes = [];
  if (quarters.length > 1) {
    notes.push(
      `These statements span ${quarters.join(' and ')}; filed under ${quarter}, `
      + 'which holds the majority of the transactions.'
    );
  }
  if (filedRows.length) {
    notes.push(
      `Added to the ${quarter} workbook already in Drive: ${filedRows.length} row`
      + `${filedRows.length === 1 ? '' : 's'} carried forward from `
      + `${filedStatements.map((x) => x.label).join(' & ')}, `
      + `${fresh.length} new row${fresh.length === 1 ? '' : 's'} from this run.`
    );
  }
  if (skipped.length) {
    notes.push(`${skipped.join(' & ')} ${skipped.length === 1 ? 'was' : 'were'} already filed and ${skipped.length === 1 ? 'was' : 'were'} not added again.`);
  }
  const wb = await buildWorkbook(ExcelJS, {
    card, quarter, transactions: merged, review, flags, statements, rules, notes,
  });
  const workbookBytes = new Uint8Array(await wb.xlsx.writeBuffer());

  onProgress('Highlighting the statements…');
  const categoryPdfs = await buildCategoryPdfs(PDFLib, sources, fresh, rules, card, quarter, basePdfs);
  const abridged = await buildAbridgedStatements(PDFLib, sources);

  return {
    card,
    cardLabel: rules.cards[card].label,
    quarter,
    fiscalYear,
    sources,
    transactions: merged,
    merged: { carried: filedRows.length, added: fresh.length, skipped },
    review,
    flags,
    summary: summarize(merged),
    parsedTotal,
    controlTotal,
    variance,
    statements,
    workbook: {
      name: rules.cards[card].workbookName.replace('{quarter}', quarter),
      bytes: workbookBytes,
    },
    categoryPdfs,
    abridged,
  };
}

/**
 * Re-run after you've filled in the COMMENTS column: apply the decisions,
 * rebuild the summary and every highlighted PDF.
 */
export async function applyReview(previous, decisions, rules, libs, onProgress = () => {}) {
  const { ExcelJS, PDFLib } = libs;
  const { rowDecisions, merchantDecisions } = decisions;

  const clean = (s) => String(s || '').replace(/ /g, '').replace(/\s+/g, ' ').trim();
  const rowKey = (d, desc, amt) => `${d}|${clean(desc)}|${round2(amt)}`;

  // "professional fees" and "Professional Fees" are the same answer
  const canon = (c) => canonicalCategory(c, rules, previous.transactions);

  const pending = new Map();
  for (const d of rowDecisions) pending.set(rowKey(d.date, d.desc, d.amount), canon(d.category));
  const byMerchant = new Map();
  for (const d of merchantDecisions) byMerchant.set(clean(d.desc).toUpperCase(), canon(d.category));

  let applied = 0;
  const moved = [];
  const reviewCat = rules.reviewCategory || 'Review';

  for (const t of previous.transactions) {
    const key = rowKey(t.date, t.desc, t.amount);
    if (t.category === reviewCat) {
      const c = pending.get(key);
      if (c) { t.category = c; t.reason = 'your review decision'; applied++; }
      continue;
    }
    const m = byMerchant.get(clean(t.desc).toUpperCase());
    if (m && m !== t.category) {
      moved.push({ desc: clean(t.desc), from: t.category, to: m, amount: t.amount });
      t.category = m;
      t.reason = 'your review decision';
      applied++;
    }
  }

  // the out-of-GTA meals rule is a standing rule, so it gets the last word here
  // too: the Review sheet asks for a category without showing where the charge
  // happened, so a meal in cottage country would otherwise stay under Meals
  const gta = applyGtaRule(previous.transactions, previous.card, rules);

  const stillReview = previous.transactions.filter((t) => t.category === reviewCat);

  onProgress('Rebuilding the workbook…');
  const notes = [
    applied
      ? `${applied} item${applied === 1 ? '' : 's'} recoded from your comments on the Review sheet.`
      : 'No new decisions were found in the COMMENTS column.',
  ];
  if (stillReview.length) {
    notes.push(`${stillReview.length} item${stillReview.length === 1 ? ' is' : 's are'} still in Review.`);
  }
  if (gta.moved.length) {
    notes.push(
      `${gta.moved.length} meal${gta.moved.length === 1 ? '' : 's'} outside the GTA `
      + `moved to ${gta.moved[0].to}: ${gta.moved.map((m) => `${m.desc} (${m.place})`).join(', ')}.`
    );
  }

  const wb = await buildWorkbook(ExcelJS, {
    card: previous.card,
    quarter: previous.quarter,
    transactions: previous.transactions,
    review: stillReview.map((t) => ({ ...t, suggested: '', note: 'Still awaiting a decision.' })),
    flags: gta.flags,
    statements: previous.statements,
    rules,
    notes,
  });

  onProgress('Rebuilding the highlighted statements…');
  const categoryPdfs = await buildCategoryPdfs(
    PDFLib, previous.sources, previous.transactions, rules, previous.card, previous.quarter
  );

  return {
    ...previous,
    applied,
    moved,
    gtaMoved: gta.moved,
    flags: gta.flags,
    review: stillReview,
    summary: summarize(previous.transactions),
    workbook: {
      name: rules.cards[previous.card].workbookName.replace('{quarter}', previous.quarter),
      bytes: new Uint8Array(await wb.xlsx.writeBuffer()),
    },
    categoryPdfs,
  };
}

function dominant(list) {
  const c = {};
  for (const x of list) c[x] = (c[x] || 0) + 1;
  return Object.keys(c).sort((a, b) => c[b] - c[a])[0];
}

function round2(n) {
  return Math.round(n * 100) / 100;
}
