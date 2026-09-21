// Orchestrates a run: statements in, categorized workbook + highlighted PDFs out.

import { readPages, quarterOf, fiscalYearOf, quarterBounds, inQuarter } from './parsers/base.js';
import { detectCard } from './parsers/registry.js';
import { categorize, summarize, applyGtaRule, canonicalCategory } from './rules.js';
import { buildWorkbook, readFiledWorkbook } from './workbook.js';
import { buildCategoryPdfs, buildAbridgedStatements } from './highlight.js';

/**
 * Read the statements and work out where they belong — nothing else. This is
 * what the Process tab calls the moment you pick files, so it can tell you which
 * card and quarter it sees, and look up what Drive already holds, BEFORE you
 * commit to a run. The parsed result is handed back so `run` can reuse it rather
 * than reading every PDF a second time.
 *
 * @param {File[]} files
 * @param {object} rules
 * @param {object} libs           {pdfjsLib}
 * @param {function} onProgress   (message) => void
 */
export async function inspect(files, rules, libs, onProgress = () => {}) {
  const { pdfjsLib } = libs;
  const sources = [];

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

  const cards = [...new Set(sources.map((s) => s.parser.id))];
  if (cards.length > 1) {
    throw new Error(
      'Those statements are from different cards ('
      + cards.map((c) => rules.cards[c].label).join(', ')
      + '). Run one card at a time so each lands in its own folder.'
    );
  }
  const card = cards[0];

  const dates = [];
  for (const s of sources) for (const t of s.parsed.transactions) dates.push(t.date);
  dates.sort();

  const quarters = [...new Set(dates.map(quarterOf))];
  const quarter = dominant(dates.map(quarterOf));
  const fiscalYear = fiscalYearOf(dates[dates.length - 1], rules.fiscalYearEndMonth);

  // how the rows fall across quarters, so the Process tab can warn you that a
  // statement straddles a boundary before you commit to anything
  const byQuarter = {};
  for (const d of dates) {
    const q = quarterOf(d);
    byQuarter[q] = (byQuarter[q] || 0) + 1;
  }

  return {
    card,
    cardLabel: rules.cards[card].label,
    quarter,
    quarters,
    byQuarter,
    fiscalYear,
    sources,
    rowCount: dates.length,
    dateRange: { from: dates[0], to: dates[dates.length - 1] },
    statements: sources.map((s) => ({
      file: s.name,
      label: s.parsed.statementLabel,
      count: s.parsed.transactions.length,
      controlTotal: s.parsed.controlTotal,
    })),
  };
}

/**
 * @param {File[]} files          statement PDFs the user dropped
 * @param {object} rules
 * @param {object} libs           {pdfjsLib, ExcelJS, PDFLib}
 * @param {function} onProgress   (message) => void
 * @param {object} opts
 *   fetchFiled   async (card, quarter, fy) => {workbookBytes, categoryPdfs} | null
 *   inspected    a prior `inspect` result, to skip re-reading the PDFs
 *   quarter      force the target quarter instead of inferring it from the dates
 *   mode         'append' (default) | 'rebuild' — 'rebuild' ignores what's filed
 */
export async function run(files, rules, libs, onProgress = () => {}, opts = {}) {
  const { ExcelJS, PDFLib } = libs;

  // ------------------------------------------------------------- 1. parse
  const seen = opts.inspected || await inspect(files, rules, libs, onProgress);
  const sources = seen.sources.slice();
  const card = seen.card;

  // ------------------------------------------- 2. group by card and quarter
  const quarters = seen.quarters;
  const quarter = opts.quarter || seen.quarter;
  const fiscalYear = seen.fiscalYear;
  const mode = opts.mode === 'rebuild' ? 'rebuild' : 'append';

  // ------------------------------------------- 2b. merge with what's filed
  // A quarter is built up over several drops. If this card and quarter already
  // have a workbook, its rows and statements come along so the new statement is
  // ADDED to the quarter rather than replacing it. De-duplication is by
  // statement, not by row: a statement already named in the filed
  // reconciliation block is skipped whole, which makes re-dropping one a no-op.
  //
  // 'rebuild' mode skips all of this: the quarter is built from the statements
  // in front of it and whatever is in Drive gets overwritten.
  let filedRows = [];
  let filedOutside = [];
  let filedStatements = [];
  let basePdfs = null;
  const skipped = [];
  if (mode === 'append' && typeof opts.fetchFiled === 'function') {
    try {
      const filed = await opts.fetchFiled(card, quarter, fiscalYear);
      if (filed && filed.workbookBytes) {
        const prior = await readFiledWorkbook(ExcelJS, filed.workbookBytes, card, rules);
        filedRows = prior.rows;
        filedOutside = prior.outside || [];
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

  const parsedFresh = [];
  for (const s of sources) {
    for (const t of s.parsed.transactions) parsedFresh.push({ ...t, source: s.name });
  }

  // ------------------------------------------- 2c. scope to the quarter
  // A statement period is not a quarter. The July statement runs from roughly
  // 18 June to 17 July, so it carries June charges that belong to Q2. Those are
  // set aside here: they stay in the workbook, on their own sheet and in the
  // reconciliation, but they are kept out of the Data sheet and therefore out
  // of every category total. The quarter's own boundaries decide this, never
  // the statement a charge happened to arrive on.
  //
  // Filed rows are re-checked too, not just new ones. A workbook built before
  // this rule existed has last quarter's spillover sitting in its Data sheet,
  // and appending the next statement quietly repairs it.
  const bounds = quarterBounds(quarter);
  const fresh = parsedFresh.filter((t) => inQuarter(t.date, quarter));
  const freshOutside = parsedFresh.filter((t) => !inQuarter(t.date, quarter));
  const carriedIn = filedRows.filter((t) => inQuarter(t.date, quarter));
  const displacedFiled = filedRows.filter((t) => !inQuarter(t.date, quarter));
  filedRows = carriedIn;

  const outside = dedupeRows([
    ...filedOutside.filter((t) => !inQuarter(t.date, quarter)),
    ...displacedFiled,
    ...freshOutside,
  ]).sort((a, b) => a.date.localeCompare(b.date) || a.desc.localeCompare(b.desc));

  if (freshOutside.length) {
    onProgress(
      `${freshOutside.length} transaction${freshOutside.length === 1 ? '' : 's'} on these `
      + `statements fall outside ${quarter} — listed separately, not counted in the summary.`
    );
  }
  if (displacedFiled.length) {
    onProgress(
      `${displacedFiled.length} row${displacedFiled.length === 1 ? '' : 's'} already filed under `
      + `${quarter} belong to another quarter — moved off the summary.`
    );
  }

  // --------------------------------------------------------- 3. categorize
  // only the new rows: anything already filed carries the category you settled
  onProgress('Categorizing…');
  const reviewCat = rules.reviewCategory || 'Review';
  const { review: freshReview, flags } = categorize(fresh, card, rules);
  categorize(freshOutside, card, rules);   // so the outside sheet shows a category too
  const merged = [...filedRows, ...fresh].sort(
    (a, b) => a.date.localeCompare(b.date) || a.desc.localeCompare(b.desc)
  );

  const carriedReview = carryForwardReview(filedRows, rules);
  const review = [...carriedReview, ...freshReview];

  // --------------------------------------------------- 4. reconcile & check
  // The summary covers the quarter, but the statements cover their own periods,
  // so the two only tie once the out-of-quarter rows are added back. Keeping
  // that identity explicit is what lets the summary be scoped to the quarter
  // without the reconciliation check losing its meaning: it still proves every
  // line on every statement was parsed and accounted for somewhere.
  const parsedTotal = round2(merged.reduce((s, t) => s + t.amount, 0));
  const outsideTotal = round2(outside.reduce((s, t) => s + t.amount, 0));
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
  const variance = round2(parsedTotal + outsideTotal - controlTotal);

  // --------------------------------------------------------- 5. build files
  onProgress('Building the workbook…');
  const notes = [];
  const overrode = opts.quarter && opts.quarter !== seen.quarter;
  if (overrode) {
    notes.push(
      `Filed under ${quarter} by hand. Left to itself the app would have chosen `
      + `${seen.quarter}, where most of these transactions fall.`
    );
  } else if (quarters.length > 1) {
    notes.push(
      `These statements span ${quarters.join(' and ')}; filed under ${quarter}, `
      + 'which holds the majority of the transactions.'
    );
  }
  if (mode === 'rebuild') {
    notes.push(
      `Built from these statements alone — anything previously filed under ${quarter} `
      + 'was replaced, not added to.'
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
  if (carriedReview.length) {
    notes.push(
      `${carriedReview.length} item${carriedReview.length === 1 ? '' : 's'} from an earlier `
      + 'statement in this quarter are still awaiting a decision and remain on the Review sheet.'
    );
  }
  if (outside.length) {
    const qs = [...new Set(outside.map((t) => quarterOf(t.date)))].sort();
    notes.push(
      `${outside.length} transaction${outside.length === 1 ? '' : 's'} totalling `
      + `$${outsideTotal.toFixed(2)} fall outside ${quarter} (${qs.join(', ')}) and are `
      + 'excluded from every category total; they are listed on the "'
      + outsideSheetName(rules, card) + '" sheet.'
    );
  }
  const wb = await buildWorkbook(ExcelJS, {
    card, quarter, transactions: merged, review, flags, statements, rules, notes,
    outside, bounds,
  });
  const workbookBytes = new Uint8Array(await wb.xlsx.writeBuffer());

  // the filed copy is renamed to the convention the card's folder already uses;
  // the file the bank gave you keeps its own name on disk
  for (const s of sources) {
    s.filedName = statementFileName(rules, card, s.parsed) || s.name;
  }

  onProgress('Highlighting the statements…');
  const categoryPdfs = await buildCategoryPdfs(PDFLib, sources, fresh, rules, card, quarter, basePdfs);
  const abridged = (await buildAbridgedStatements(PDFLib, sources)).map((a, i) => ({
    ...a,
    name: sources[i] ? sources[i].filedName : a.name,
  }));

  return {
    card,
    cardLabel: rules.cards[card].label,
    quarter,
    quarterAuto: seen.quarter,
    quarterOverridden: !!overrode,
    mode,
    fiscalYear,
    sources,
    transactions: merged,
    merged: {
      carried: filedRows.length,
      added: fresh.length,
      skipped,
      carriedReview: carriedReview.length,
      carriedFx: filedRows.filter((t) => t.fx).length,
    },
    review,
    flags,
    summary: summarize(merged),
    parsedTotal,
    outside,
    outsideTotal,
    bounds,
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
    // the out-of-quarter list is untouched by a review pass, but it has to be
    // handed back in or the rebuilt workbook would lose the sheet entirely
    outside: previous.outside || [],
    bounds: previous.bounds,
    transactions: previous.transactions,
    review: stillReview.map((t) => ({
      ...t,
      suggested: t.suggested || '',
      // keep whatever context the row already carried — on Rogers the note is
      // the only record of the town, and losing it makes the row unanswerable
      note: t.note ? `${t.note} Still awaiting a decision.` : 'Still awaiting a decision.',
    })),
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

/**
 * A filed row still sitting in Review is a question you haven't answered yet, so
 * it has to go back onto the Review sheet of the rebuilt workbook.
 *
 * Without this, adding next month's statement quietly buries it: the row stays
 * coded Review on the data sheet and keeps dragging the summary's Review bucket
 * up, but nothing on the Review sheet asks you about it any more, so it can sit
 * there for the rest of the quarter. Its note and suggestion are read back out
 * of the filed workbook rather than regenerated — on Rogers that note is the
 * only record of which town the charge happened in, and without it the row
 * cannot be answered.
 */
export function carryForwardReview(filedRows, rules) {
  const reviewCat = rules.reviewCategory || 'Review';
  return filedRows
    .filter((t) => t.category === reviewCat)
    .map((t) => ({
      ...t,
      suggested: t.suggested || '',
      note: t.note || 'Carried forward from an earlier statement in this quarter.',
    }));
}

const MON = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/**
 * What a statement PDF should be called once it's filed.
 *
 * Banks name their downloads unhelpfully and inconsistently, so the file is
 * renamed to the convention each card's folder already uses. The template lives
 * in rules.json per card; the defaults here match what's in the audit folder
 * today, so an existing rules file without the key still gets the right name.
 *
 *   {card}  the card's label        Amex
 *   {mon}   three-letter month      Jul
 *   {day}   day of month, no pad    17
 *   {year}  four-digit year         2026
 */
export function statementFileName(rules, card, parsed) {
  const cfg = rules.cards[card] || {};
  const tpl = cfg.statementFileName || DEFAULT_STATEMENT_NAMES[card]
    || '{card} {mon} {day} {year} Statement.pdf';
  const iso = parsed.statementDate;
  if (!iso) return null;
  const [y, m, d] = iso.split('-').map(Number);
  return tpl
    .replace('{card}', cfg.label || card)
    .replace('{mon}', MON[m])
    .replace('{day}', String(d))
    .replace('{year}', String(y));
}

const DEFAULT_STATEMENT_NAMES = {
  amex: '{card} {mon} {day} {year} Statement.pdf',
  cibc: '{card} Statement - {mon} {day} {year}.pdf',
  rogers: '{card} {mon} {day} {year} Statement.pdf',
};

/** The sheet that holds transactions belonging to a different quarter. */
export function outsideSheetName(rules, card) {
  return ((rules.cards[card] || {}).sheets || {}).outside || 'Outside This Quarter';
}

/**
 * Drop repeats of the same transaction. The out-of-quarter list is rebuilt from
 * several places at once — what the filed workbook already held, rows displaced
 * out of its Data sheet, and rows off the statements in front of us — and the
 * same charge can legitimately reach it by more than one route.
 */
function dedupeRows(rows) {
  const seen = new Set();
  const out = [];
  for (const t of rows) {
    const key = `${t.date}|${String(t.desc || '').replace(/\s+/g, ' ').trim().toUpperCase()}`
      + `|${round2(t.amount)}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(t);
  }
  return out;
}

function dominant(list) {
  const c = {};
  for (const x of list) c[x] = (c[x] || 0) + 1;
  return Object.keys(c).sort((a, b) => c[b] - c[a])[0];
}

function round2(n) {
  return Math.round(n * 100) / 100;
}
