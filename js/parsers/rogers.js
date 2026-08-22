// Rogers Bank Red World Elite Mastercard statement parser.
//
// Layout: "Jun 19 | Jun 22 | MERCHANT | CITY | PROV | 33.90", where the three
// description fields are fixed-width and space-padded. The font is proportional,
// so those fields do NOT land on fixed x positions — but pdf.js emits each field
// as its own text item, so we split them on the field index base.js records
// rather than guessing from gaps. That matters for rows like
// "A&W #4514 TORONTO   TORONTO   ON", where the merchant name itself ends in the
// city name and no gap rule could tell the two apart.
//
// Reconciliation is two-sided, which is stronger than what the other cards allow:
// the statement prints "New purchases & debits" and "Payments & credits", and the
// positive and negative rows have to tie to those two figures exactly. Card
// payments are dropped (they are not an expense); refunds and rebates are kept as
// negative rows so they reduce the category they belong to, and the control total
// is adjusted by the payments we removed so the check still bites.

import { MONTHS, AMOUNT_RE, rowsOf, toNumber, resolveYear, isoDate } from './base.js';

export const id = 'rogers';
export const label = 'Rogers';

const PERIOD_RE = /Statement Period\s+([A-Z][a-z]{2})\s+(\d{1,2}),\s+(\d{4})\s*-\s*([A-Z][a-z]{2})\s+(\d{1,2}),\s+(\d{4})/;
const PURCHASES_RE = /New purchases & debits\s+\$?(-?[\d,]+\.\d{2})/;
const CREDITS_RE = /Payments & credits\s+\$?(-?[\d,]+\.\d{2})/;
const FX_RE = /FOREIGN CURRENCY\s+([A-Z]{3})\s+([\d,]+\.?\d*)\s*@\s*([\d.]+)/;
const PAYMENT_RE = /^PAYMENT[, ]|^PRE-?AUTHORIZED PAYMENT|THANK YOU$/i;

// Trailing 2-letter codes get glued to a city that fills its 13-character field
// exactly: "SAN FRANCISCOCA", "RICHMOND HILLON".
const REGIONS = [
  'ON', 'QC', 'NL', 'NS', 'NB', 'PE', 'MB', 'SK', 'AB', 'BC', 'YT', 'NT', 'NU',
  'AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA', 'HI', 'ID', 'IL',
  'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD', 'MA', 'MI', 'MN', 'MS', 'MO', 'MT',
  'NE', 'NV', 'NH', 'NJ', 'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI',
  'SC', 'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY', 'DC',
];
const GLUED_RE = new RegExp(`^(.*[A-Za-z])(${REGIONS.join('|')})$`);

export function detect(pages) {
  const head = pages.slice(0, 2).map((p) => p.words.map((w) => w.text).join(' ')).join(' ');
  return /rogersbank\.com/i.test(head) || /Rogers\s+Bank/i.test(head);
}

function statementPeriod(pages) {
  for (const page of pages) {
    for (const row of rowsOf(page)) {
      const m = row.text.match(PERIOD_RE);
      if (m) return { endMonth: MONTHS[m[4]], endDay: +m[5], year: +m[6] };
    }
  }
  return null;
}

/** The two figures the transactions have to tie back to. */
function controlFigures(pages) {
  let purchases = null;
  let credits = null;
  for (const page of pages) {
    for (const row of rowsOf(page)) {
      if (purchases === null) {
        const m = row.text.match(PURCHASES_RE);
        if (m) purchases = toNumber(m[1]);
      }
      if (credits === null) {
        const m = row.text.match(CREDITS_RE);
        if (m) credits = toNumber(m[1]);
      }
    }
  }
  return { purchases, credits };
}

/** x boundaries from this page's "Date Date Description Amount ($)" header. */
function columnsOf(rows) {
  for (const row of rows) {
    const txt = row.words.map((w) => w.text);
    const dates = txt.filter((t) => t === 'Date').length;
    if (dates === 2 && txt.includes('Description') && txt.includes('Amount')) {
      const post = row.words.filter((w) => w.text === 'Date')[1];
      const desc = row.words.find((w) => w.text === 'Description');
      const amt = row.words.find((w) => w.text === 'Amount');
      return {
        post: post.x0 - 6,
        desc: desc.x0 - 8,
        amt: amt.x0 - 8,
        span: [row.x0 - 6, row.x1 + 6],
      };
    }
  }
  return null;
}

/**
 * Split the description column into merchant / place using the field index.
 * Field 1 is the merchant; anything after it is city (+ province or state).
 */
function splitDescription(words) {
  const groups = [];
  for (const w of words) {
    const last = groups[groups.length - 1];
    if (last && last.field === w.field) last.words.push(w.text);
    else groups.push({ field: w.field, words: [w.text] });
  }
  const merchant = groups.length ? groups[0].words.join(' ') : '';
  const rest = groups.slice(1).map((g) => g.words.join(' '));

  if (!rest.length) return { merchant, place: '' };

  // last group is the province/state unless the city filled its 13-character
  // field, in which case the two ran together into one ("SAN FRANCISCOCA")
  let city = rest.slice(0, -1).join(' ');
  let region = rest[rest.length - 1];
  if (!city) {
    const tail = region.split(/\s+/);
    const glued = region.length > 2 && region.match(GLUED_RE);
    if (tail.length > 1 && REGIONS.includes(tail[tail.length - 1])) {
      region = tail.pop();
      city = tail.join(' ');
    } else if (glued) {
      city = glued[1].trim();
      region = glued[2];
    } else {
      city = region;
      region = '';
    }
  }
  return { merchant, place: region ? `${city}, ${region}` : city };
}

export function parse(pages) {
  const period = statementPeriod(pages);
  if (!period) throw new Error('Could not find the statement period on this Rogers PDF.');
  const { purchases, credits } = controlFigures(pages);

  const txns = [];
  const payments = [];
  const spans = {};

  pages.forEach((page, pno) => {
    const rows = rowsOf(page);
    const cols = columnsOf(rows);
    if (!cols) return;
    spans[pno] = cols.span;

    for (const row of rows) {
      const trans = row.words.filter((w) => w.x0 < cols.post);
      const desc = row.words.filter((w) => w.x0 >= cols.desc && w.x0 < cols.amt);
      const amt = row.words.filter((w) => w.x0 >= cols.amt);

      // "FOREIGN CURRENCY  USD  5.67 @ 1.455026455" sits under its transaction
      if (!trans.length && !amt.length && desc.length) {
        const m = row.text.match(FX_RE);
        if (m && txns.length) {
          const last = txns[txns.length - 1];
          last.fx = { currency: m[1], amount: toNumber(m[2]), rate: parseFloat(m[3]) };
          last.y1 = Math.max(last.y1, row.y1);
        }
        continue;
      }

      if (trans.length < 2 || !amt.length || !desc.length) continue;
      const month = MONTHS[trans[0].text];
      if (!month || !/^\d{1,2}$/.test(trans[1].text)) continue;
      if (!AMOUNT_RE.test(amt[amt.length - 1].text)) continue;

      const { merchant, place } = splitDescription(desc);
      if (!merchant) continue;

      const entry = {
        date: isoDate(resolveYear(month, period.year, period.endMonth), month, +trans[1].text),
        desc: merchant,
        place,
        spendCategory: '',
        amount: toNumber(amt[amt.length - 1].text),
        fx: null,
        detail: [],
        page: pno,
        y0: row.y0,
        y1: row.y1,
      };

      // a payment of the card balance is not an expense
      if (PAYMENT_RE.test(merchant)) payments.push(entry);
      else txns.push(entry);
    }
  });

  // Control = purchases, less the credits we kept. The statement lumps payments
  // in with credits, so add back what we dropped; a credit row the parser missed
  // still shows up as a variance.
  const paid = payments.reduce((s, p) => s + p.amount, 0);
  const control = purchases === null || credits === null
    ? null
    : Math.round((purchases - credits - paid) * 100) / 100;

  const names = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return {
    card: id,
    statementLabel: `${names[period.endMonth]} ${period.endDay}, ${period.year}`,
    statementDate: isoDate(period.year, period.endMonth, period.endDay),
    controlTotal: control,
    transactions: txns,
    transactionPages: [...new Set(txns.map((t) => t.page))].sort((a, b) => a - b),
    _spans: spans,
    _excludedPayments: payments.length,
  };
}

// Rogers rows are 11.3pt apart in a 7.7pt font, so the generic box is taller
// than the gap. Shave it back to sit on the line it belongs to.
export const highlightInset = 2.3;

export function highlightSpan(page, parsed, pageNo) {
  const s = parsed && parsed._spans && parsed._spans[pageNo];
  return s || [20, page.width - 20];
}
