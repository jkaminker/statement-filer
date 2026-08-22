// Produces one PDF per expense category: the statement's transaction pages with
// that category's line items highlighted. Uses real PDF Highlight annotations —
// the same thing the files already in the audit folder use — so the text stays
// crisp and readable underneath, and the highlight can be clicked or removed in
// any PDF reader.

const PAD = 1.6;

export async function buildCategoryPdfs(PDFLib, sources, transactions, rules, card, quarter) {
  const { PDFDocument, PDFString } = PDFLib;
  const cfg = rules.cards[card];
  const [cr, cg, cb] = rules.highlightColor || [0.988, 0.957, 0.522];

  const byCategory = {};
  for (const t of transactions) {
    if (t.page === undefined || t.page === null) continue;
    (byCategory[t.category] = byCategory[t.category] || []).push(t);
  }

  const out = [];
  for (const category of Object.keys(byCategory).sort()) {
    const rows = byCategory[category];
    const doc = await PDFDocument.create();

    // keep only the transaction pages, in statement order, and remember where
    // each original page landed in the combined document
    const remap = new Map();
    for (const src of sources) {
      const donor = await PDFDocument.load(src.bytes);
      const wanted = src.parsed.transactionPages;
      const copied = await doc.copyPages(donor, wanted);
      copied.forEach((p, i) => {
        doc.addPage(p);
        remap.set(`${src.name}::${wanted[i]}`, doc.getPageCount() - 1);
      });
    }

    let lines = 0;
    for (const t of rows) {
      const idx = remap.get(`${t.source}::${t.page}`);
      if (idx === undefined) continue;
      const page = doc.getPage(idx);
      const { width, height } = page.getSize();
      const src = sources.find((s) => s.name === t.source);
      const [x0, x1] = spanFor(src, t.page, width);

      // parsed y is measured from the top of the page; PDF space measures from
      // the bottom, so flip before writing the annotation rectangle
      const yTop = height - t.y0 + PAD;
      const yBottom = height - t.y1 - PAD;

      const annot = doc.context.obj({
        Type: 'Annot',
        Subtype: 'Highlight',
        Rect: [x0, yBottom, x1, yTop],
        QuadPoints: [x0, yTop, x1, yTop, x0, yBottom, x1, yBottom],
        C: [cr, cg, cb],
        CA: 1,
        F: 4, // print
        T: PDFString.of('Category'),
        Contents: PDFString.of(category),
      });
      page.node.addAnnot(doc.context.register(annot));
      lines++;
    }

    const bytes = await doc.save();
    out.push({
      category,
      fileName: cfg.categoryFileName
        .replace('{quarter}', quarter)
        .replace('{category}', category),
      bytes,
      lines,
      total: Math.round(rows.reduce((s, t) => s + t.amount, 0) * 100) / 100,
    });
  }
  return out;
}

function spanFor(src, pageNo, width) {
  const parser = src.parser;
  if (parser && typeof parser.highlightSpan === 'function') {
    const meta = src.pagesMeta[pageNo] || { width };
    const span = parser.highlightSpan(meta, src.parsed, pageNo);
    if (span && span.length === 2 && span[1] > span[0]) return span;
  }
  return [22.64, width - 22.64];
}
