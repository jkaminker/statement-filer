// American Express (Aeroplan Reserve) statement parser.
//
// Layout: "Jun 17  Jun 18  MERCHANT  CITY   80.34", with an optional indented
// follow-on line carrying the foreign-currency detail ("UNITED STATES DOLLAR 55.90 @ 1.43721")
// or flight routing / hotel stay detail. The control total is "Total of New Transactions".

import { MONTHS, AMOUNT_RE, rowsOf, toNumber, resolveYear, isoDate } from './base.js';

const FX_RE = /^(.+?)\s+([\d,]+\.\d+)\s+@\s+([\d.]+)$/;
const CLOSING_RE = /Closing Date/i;
const DATE_RE = /([A-Z][a-z]{2})\s+(\d{1,2}),\s+(\d{4})/;

export const id = 'amex';
export const label = 'Amex';

export function detect(pages) {
  const head = pages.slice(0, 2).map((p) => p.words.map((w) => w.text).join(' ')).join(' ');
  return /American Express/i.test(head) && /Statement of Account/i.test(head);
}

/** Closing date drives the year and the statement label. */
function statementDate(pages) {
  for (const page of pages.slice(0, 2)) {
    const rows = rowsOf(page);
    for (let i = 0; i < rows.length; i++) {
      if (!CLOSING_RE.test(rows[i].text)) continue;
      for (let j = i; j < Math.min(i + 4, rows.length); j++) {
        const all = [...rows[j].text.matchAll(new RegExp(DATE_RE, 'g'))];
        if (all.length) {
          const m = all[all.length - 1]; // closing date is the last of opening/closing
          return { month: MONTHS[m[1]], day: +m[2], year: +m[3] };
        }
      }
    }
  }
  return null;
}

export function parse(pages) {
  const stmt = statementDate(pages);
  if (!stmt) throw new Error('Could not find the statement closing date on this Amex PDF.');

  const txns = [];
  let control = null;
  // The "New Transactions for …" header appears once, on the first transaction
  // page; continuation pages just carry the column headings. So this flag has to
  // survive page boundaries and only close on an explicit end marker.
  let live = false;

  pages.forEach((page, pno) => {
    for (const row of rowsOf(page)) {
      const t = row.text;

      if (/^New Transactions for/i.test(t)) { live = true; continue; }
      if (/^Total of New Transactions/i.test(t)) {
        const m = t.match(/(-?[\d,]+\.\d{2})\s*$/);
        if (m) control = toNumber(m[1]);
        live = false;
        continue;
      }
      if (/^(New Payments|Total of Payment Activity|Charges Made in Foreign)/i.test(t)) {
        live = false;
        continue;
      }
      if (!live) continue;

      const parts = t.split(/\s+/);
      const looksLikeTxn =
        parts.length >= 5 &&
        MONTHS[parts[0]] &&
        /^\d{1,2}$/.test(parts[1]) &&
        MONTHS[parts[2].replace('*', '')] &&
        /^\d{1,2}$/.test(parts[3]) &&
        AMOUNT_RE.test(parts[parts.length - 1]);

      if (looksLikeTxn) {
        const month = MONTHS[parts[0]];
        const desc = parts.slice(4, -1).join(' ');
        if (!desc) continue;
        txns.push({
          date: isoDate(resolveYear(month, stmt.year, stmt.month), month, +parts[1]),
          desc,
          amount: toNumber(parts[parts.length - 1]),
          fx: null,
          detail: [],
          page: pno,
          y0: row.y0,
          y1: row.y1,
          x0: row.x0,
          x1: row.x1,
        });
        continue;
      }

      // indented continuation line belonging to the transaction above
      if (txns.length && !MONTHS[parts[0]]) {
        const last = txns[txns.length - 1];
        if (last.page !== pno) continue;
        last.detail.push(t);
        const fx = t.match(FX_RE);
        if (fx && /DOLLAR|EURO|POUND|YEN|FRANC|PESO|SHEKEL|KRONA/i.test(fx[1])) {
          last.fx = { currency: fx[1].trim(), amount: toNumber(fx[2]), rate: parseFloat(fx[3]) };
        }
      }
    }
  });

  const label = `${['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][stmt.month]} ${stmt.day}, ${stmt.year}`;
  return {
    card: id,
    statementLabel: label,
    statementDate: isoDate(stmt.year, stmt.month, stmt.day),
    controlTotal: control,
    transactions: txns,
    // pages that carry at least one transaction - these are the ones we highlight
    transactionPages: [...new Set(txns.map((t) => t.page))].sort((a, b) => a - b),
  };
}

/** Full-width highlight bar geometry, matching the existing Q1/Q2 files. */
export function highlightSpan(page) {
  return [22.64, page.width - 22.64];
}
