// CIBC Costco World Mastercard statement parser.
//
// Layout: "Jul 15  Jul 15  [bonus mark]  MERCHANT  CITY  PROV | Spend Category | 248.02".
// Column boundaries come from each page's own header row, because continuation pages
// use a different left margin. A rotated barcode watermark (*0202560000*) lands inside
// transaction rows and is stripped in base.rowsOf.
//
// Control total is "Total for <card number>", which is purchases net of credits.

import { MONTHS, AMOUNT_RE, rowsOf, toNumber, resolveYear, isoDate } from './base.js';

const FX_RE = /^([\d,.]+)\s+([A-Z]{3})\s+@\s+([\d.]+)/;
const PROVINCES = ['ON', 'QC', 'NL', 'NS', 'NB', 'PE', 'MB', 'SK', 'AB', 'BC', 'YT', 'NT', 'NU'];
const GLUED_RE = new RegExp(`^(.*[A-Za-z])(${PROVINCES.join('|')})$`);
const PERIOD_RE = /Transactions from ([A-Z][a-z]+) (\d{1,2}) to ([A-Z][a-z]+) (\d{1,2}), (\d{4})/;

export const id = 'cibc';
export const label = 'CIBC';

export function detect(pages) {
  const head = pages.slice(0, 2).map((p) => p.words.map((w) => w.text).join(' ')).join(' ');
  return /CIBC/i.test(head) && /Mastercard/i.test(head);
}

function statementPeriod(pages) {
  for (const page of pages) {
    for (const row of rowsOf(page)) {
      const m = row.text.match(PERIOD_RE);
      if (m) {
        return {
          endMonth: MONTHS[m[3]],
          endDay: +m[4],
          year: +m[5],
        };
      }
    }
  }
  return null;
}

/** x boundaries from the "date date Description Spend Categories Amount($)" header. */
function columnsOf(rows) {
  for (const row of rows) {
    const txt = row.words.map((w) => w.text);
    const dates = txt.filter((t) => t === 'date').length;
    if (dates === 2 && txt.includes('Description') && txt.includes('Spend') && txt.includes('Amount($)')) {
      const xs = {};
      row.words.forEach((w) => { if (!(w.text in xs)) xs[w.text] = w.x0; });
      const postX = row.words.filter((w) => w.text === 'date')[1].x0;
      return {
        post: postX - 6,
        desc: xs.Description - 8,
        cat: xs.Spend - 8,
        amt: xs['Amount($)'] - 8,
        span: [row.x0 - 4, row.x1 + 4],
      };
    }
  }
  return null;
}

/** "MERCHANT CITY PROV" -> " MERCHANT CITY, PROV", matching the existing workbooks. */
function formatDesc(words) {
  const w = words.slice();
  if (w.length >= 2 && w[w.length - 1].length > 2) {
    const m = w[w.length - 1].match(GLUED_RE); // "ROCKY HARBOURNL"
    if (m) w.splice(w.length - 1, 1, m[1], m[2]);
  }
  return ' ' + w.slice(0, -1).join(' ') + ', ' + w[w.length - 1];
}

export function parse(pages) {
  const period = statementPeriod(pages);
  if (!period) throw new Error('Could not find the statement period on this CIBC PDF.');

  const txns = [];
  const spans = {};
  let control = null;

  pages.forEach((page, pno) => {
    const rows = rowsOf(page);
    const cols = columnsOf(rows);
    if (!cols) return;
    spans[pno] = cols.span;

    let live = false;
    for (const row of rows) {
      const t = row.text;
      if (/^Your new charges and credits/i.test(t)) { live = true; continue; }
      if (/^Total for /i.test(t)) {
        const m = t.match(/(-?[\d,]+\.\d{2})\s*$/);
        if (m) control = toNumber(m[1]);
        live = false;
        continue;
      }
      if (/^(Information about your|Your message centre|Your payments|Your interest|Total payments)/i.test(t)) {
        live = false;
        continue;
      }
      if (!live) continue;

      const trans = row.words.filter((w) => w.x0 < cols.post);
      const desc = row.words.filter((w) => w.x0 >= cols.desc && w.x0 < cols.cat);
      const cat = row.words.filter((w) => w.x0 >= cols.cat && w.x0 < cols.amt);
      const amt = row.words.filter((w) => w.x0 >= cols.amt);

      // foreign-currency continuation line
      if (txns.length && !trans.length && !amt.length && desc.length) {
        const m = desc.map((w) => w.text).join(' ').match(FX_RE);
        if (m) {
          const last = txns[txns.length - 1];
          last.fx = { amount: toNumber(m[1]), currency: m[2], rate: parseFloat(m[3]) };
          last.y1 = Math.max(last.y1, row.y1);
        }
        continue;
      }

      if (trans.length < 2 || !amt.length) continue;
      if (!AMOUNT_RE.test(amt[amt.length - 1].text)) continue;
      const month = MONTHS[trans[0].text];
      if (!month || !/^\d{1,2}$/.test(trans[1].text)) continue;

      // the bonus-rewards mark sometimes floats into the description column
      const dwords = desc.map((w) => w.text).filter((s) => s !== 'Ý' && s !== 'Ý');
      if (!dwords.length) continue;

      txns.push({
        date: isoDate(resolveYear(month, period.year, period.endMonth), month, +trans[1].text),
        desc: formatDesc(dwords),
        spendCategory: cat.map((w) => w.text).join(' '),
        amount: toNumber(amt[amt.length - 1].text),
        fx: null,
        detail: [],
        page: pno,
        y0: row.y0,
        y1: row.y1,
      });
    }
  });

  const names = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return {
    card: id,
    statementLabel: `${names[period.endMonth]} ${period.endDay}, ${period.year}`,
    statementDate: isoDate(period.year, period.endMonth, period.endDay),
    controlTotal: control,
    transactions: txns,
    transactionPages: [...new Set(txns.map((t) => t.page))].sort((a, b) => a - b),
    _spans: spans,
  };
}

export function highlightSpan(page, parsed, pageNo) {
  const s = parsed && parsed._spans && parsed._spans[pageNo];
  return s || [22.64, page.width - 22.64];
}
