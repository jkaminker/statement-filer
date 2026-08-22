// Shared PDF reading helpers. Everything runs in the browser via pdf.js.
//
// The statements we parse are laid out as fixed columns, so the reliable unit is
// not "a text item" but "a visual row": every glyph run sharing a baseline, sorted
// left to right. Column boundaries are then read off each page's own header row,
// because continuation pages sit at a different left margin than the first page.

export const MONTHS = {
  Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6,
  Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12,
};

export const AMOUNT_RE = /^-?[\d,]+\.\d{2}$/;
const WATERMARK_RE = /^\*\d+\*$/;

/**
 * Pull every page's words with positions.
 * @returns {Promise<Array<{width:number,height:number,words:Array}>>}
 *   word = {text, x0, y0, x1, y1}  (y measured from the TOP of the page)
 */
export async function readPages(pdfjsLib, arrayBuffer) {
  const doc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const pages = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const viewport = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();
    const words = [];
    // Some statements (Rogers) lay the description out as fixed-width fields —
    // merchant, city, province — padded out with spaces. pdf.js emits each field
    // as its own text item, so remembering which item a word came from is what
    // lets a parser split "A&W #4514 TORONTO" (the merchant) from "TORONTO" (the
    // city) in a case where the gap in x cannot.
    let field = 0;
    for (const item of content.items) {
      const text = (item.str || '').trim();
      if (!text) continue;
      field++;
      // pdf.js transform: [a,b,c,d,e,f]; e,f are the origin in PDF space.
      // Group rows by BASELINE, not by glyph top: a row mixes font sizes (the
      // amount column is bolder than the description) and tops therefore differ
      // by several points on what is visually one line.
      const x0 = item.transform[4];
      const baseline = item.transform[5];
      const h = item.height || Math.abs(item.transform[3]) || 8;
      const yTop = viewport.height - baseline - h;          // measured from page top
      const yBottom = viewport.height - baseline + h * 0.25; // room for descenders
      // one text item can hold several space-separated words; split so that
      // column boundaries land between words, not through the middle of one
      const parts = text.split(/\s+/).filter(Boolean);
      const per = (item.width || 0) / Math.max(text.length, 1);
      let cursor = x0;
      for (const part of parts) {
        const w = per * part.length;
        words.push({
          text: part, x0: cursor, x1: cursor + w,
          y0: yTop, y1: yBottom, baseline, field,
        });
        cursor += w + per; // approximate the space
      }
    }
    pages.push({ width: viewport.width, height: viewport.height, words });
  }
  return { doc, pages };
}

/** Group a page's words into visual rows, top to bottom, each sorted left to right. */
export function rowsOf(page, tol = 2.5) {
  const words = page.words
    .filter((w) => !WATERMARK_RE.test(w.text))
    .slice()
    .sort((a, b) => b.baseline - a.baseline || a.x0 - b.x0); // top of page first
  const rows = [];
  let cur = [];
  for (const w of words) {
    if (cur.length && Math.abs(w.baseline - cur[0].baseline) > tol) {
      rows.push(cur);
      cur = [];
    }
    cur.push(w);
  }
  if (cur.length) rows.push(cur);
  return rows.map((r) => {
    r.sort((a, b) => a.x0 - b.x0);
    return {
      words: r,
      text: r.map((w) => w.text).join(' '),
      y0: Math.min(...r.map((w) => w.y0)),
      y1: Math.max(...r.map((w) => w.y1)),
      x0: Math.min(...r.map((w) => w.x0)),
      x1: Math.max(...r.map((w) => w.x1)),
    };
  });
}

export function toNumber(s) {
  return parseFloat(String(s).replace(/,/g, ''));
}

/** Statement years roll over in December; a Dec row on a Jan statement is last year. */
export function resolveYear(month, statementYear, statementMonth) {
  if (month === 12 && statementMonth < 12) return statementYear - 1;
  if (month === 1 && statementMonth === 12) return statementYear + 1;
  return statementYear;
}

export function isoDate(y, m, d) {
  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

/** Calendar quarter label used for the folder names, e.g. "Q3 2026". */
export function quarterOf(isoStr) {
  const [y, m] = isoStr.split('-').map(Number);
  return `Q${Math.floor((m - 1) / 3) + 1} ${y}`;
}

/** Fiscal year that a date falls in, given a fiscal year-end month. */
export function fiscalYearOf(isoStr, fyEndMonth) {
  const [y, m] = isoStr.split('-').map(Number);
  return m > fyEndMonth ? y + 1 : y;
}
