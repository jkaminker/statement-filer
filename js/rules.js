// The categorization engine. Every decision it makes is traceable to a line in
// rules.json, and anything it cannot settle goes to the Review sheet rather than
// being guessed silently.

const LS_KEY = 'statement-filer.rules';

export async function loadRules() {
  const stored = localStorage.getItem(LS_KEY);
  if (stored) {
    try { return JSON.parse(stored); } catch (_) { /* fall through to the file */ }
  }
  const res = await fetch('rules.json', { cache: 'no-store' });
  if (!res.ok) throw new Error('Could not load rules.json');
  return res.json();
}

export function saveRules(rules) {
  localStorage.setItem(LS_KEY, JSON.stringify(rules));
}

export function resetRules() {
  localStorage.removeItem(LS_KEY);
}

const clean = (s) => String(s || '').replace(/ /g, '').replace(/\s+/g, ' ').trim();

/** Merchant key: the description with the trailing ", PROV" stripped (CIBC style). */
function merchantKey(desc) {
  return clean(desc);
}

function longestMatch(place, list) {
  let best = null;
  for (const token of list) {
    if (place.includes(token) && (!best || token.length > best.length)) best = token;
  }
  return best;
}

/**
 * You often shop the same place on more than one card. When a merchant is new to
 * THIS card but you have already ruled on it elsewhere, offer that answer as the
 * suggestion on the Review sheet — and only when every other card agrees. It is
 * never applied on its own; the row still waits for you.
 *
 * Cards differ in how much of the line they keep: Rogers stores "SDM 1281" while
 * CIBC stores "SDM 1281 TORONTO, ON", so match on the leading merchant text.
 */
function crossCardSuggestion(key, cardId, rules) {
  const target = key.toUpperCase();
  if (target.length < 4) return null;
  const cats = new Set();
  const cards = [];
  for (const otherId of Object.keys(rules.merchants || {})) {
    if (otherId === cardId || otherId === 'shared') continue;
    for (const k of Object.keys(rules.merchants[otherId])) {
      const kk = clean(k).toUpperCase();
      if (kk === target || kk.startsWith(target + ' ')) {
        cats.add(rules.merchants[otherId][k]);
        if (!cards.includes(otherId)) cards.push(otherId);
      }
    }
  }
  if (cats.size !== 1) return null;
  return { category: [...cats][0], cards };
}

/**
 * Where a charge happened: 'gta' | 'outside' | 'artifact' | null (unknown).
 * Amex and CIBC carry the town inside the description ("… TORONTO, ON"); Rogers
 * keeps it in its own field, so a parser can pass it separately as `where`.
 */
export function classifyPlace(desc, gtaRule, where) {
  const d = clean(desc).toUpperCase();
  for (const k of Object.keys(gtaRule.addressArtifacts || {})) {
    if (d.includes(k)) return 'artifact';
  }
  const src = where ? clean(where).toUpperCase() : d;
  const place = src.includes(', ') ? src.slice(0, src.lastIndexOf(', ')) : src;
  const out = longestMatch(place, gtaRule.outsidePlaces || []);
  const gta = longestMatch(place, gtaRule.gtaPlaces || []);
  if (out && gta) return out.length >= gta.length ? 'outside' : 'gta';
  if (out) return 'outside';
  if (gta) return 'gta';
  return null;
}

/**
 * Categorize a parsed statement's transactions.
 * @returns {{transactions:Array, review:Array, flags:Array}}
 *   each transaction gains {category, reason}
 */
export function categorize(transactions, cardId, rules) {
  const merchants = { ...(rules.merchants.shared || {}), ...(rules.merchants[cardId] || {}) };
  const alwaysReview = (rules.alwaysReview || []).filter((r) => !r.card || r.card === cardId);
  const REVIEW = rules.reviewCategory || 'Review';
  const review = [];
  const flags = [];
  const seenAmounts = {};

  for (const t of transactions) {
    const key = merchantKey(t.desc);
    const upper = key.toUpperCase();

    // 1. always-review entries. A plain entry only bites when you haven't already
    //    ruled on that merchant — once a decision is in the merchant map it wins,
    //    so the app never re-asks a question you've answered. Set "force": true on
    //    an entry to have it flagged every quarter regardless.
    const forced = alwaysReview.find(
      (r) => upper.includes(r.match.toUpperCase()) && (r.force || !merchants[key])
    );
    if (forced) {
      t.category = REVIEW;
      t.reason = 'flagged in rules';
      review.push({
        ...t,
        suggested: forced.suggested || merchants[key] || '',
        note: forced.note || '',
      });
      continue;
    }

    // 2. exact merchant match from your prior workbooks
    if (merchants[key]) {
      t.category = merchants[key];
      t.reason = 'merchant map';
    } else {
      // 3. the bank's own spend category, mapped the way your files map it
      const fallback = (rules.spendCategoryDefaults || {})[t.spendCategory];
      if (fallback && fallback !== REVIEW) {
        t.category = fallback;
        t.reason = `spend category "${t.spendCategory}"`;
      } else {
        t.category = REVIEW;
        t.reason = 'new merchant, no rule';
        const cross = crossCardSuggestion(key, cardId, rules);
        const note = t.spendCategory
          ? `New merchant on this card. Bank category: ${t.spendCategory}.`
          : 'New merchant on this card - no prior treatment to follow.';
        review.push({
          ...t,
          suggested: (fallback && fallback !== REVIEW ? fallback : '') || (cross ? cross.category : ''),
          note: cross
            ? `${note} You code this merchant as "${cross.category}" on your `
              + `${cross.cards.map((c) => (rules.cards[c] || {}).label || c).join(' and ')} `
              + 'statements - leave COMMENTS blank and it stays in Review, or confirm it there.'
            : note,
        });
        continue;
      }
    }

    // 4. large first-time charges are worth a look even when a rule matched
    const lar = rules.largeAmountReview;
    if (lar && lar.enabled && !merchants[key] && t.amount >= lar.threshold) {
      t.category = REVIEW;
      t.reason = `new merchant over $${lar.threshold}`;
      review.push({
        ...t,
        suggested: t.category,
        note: `First time seeing this merchant and the charge is $${t.amount.toFixed(2)}.`,
      });
      continue;
    }
    seenAmounts[key] = (seenAmounts[key] || 0) + t.amount;
  }

  // 5. the out-of-GTA meals rule, applied after everything else
  const g = rules.gtaRule;
  if (g && g.enabled) {
    // your Rogers workbooks call this category "Travel" where Amex and CIBC call
    // it "Business Travel", so each card can name its own
    const card = rules.cards[cardId] || {};
    const travelCategory = card.gtaToCategory || g.toCategory;
    for (const t of transactions) {
      if (!(g.fromCategories || []).includes(t.category)) continue;
      const where = classifyPlace(t.desc, g, t.place);
      if (where === 'outside') {
        t.category = travelCategory;
        t.reason = 'meal outside the GTA';
      } else if (where === 'artifact') {
        flags.push({
          desc: clean(t.desc),
          category: t.category,
          amount: t.amount,
          note: g.addressArtifacts[
            Object.keys(g.addressArtifacts).find((k) => clean(t.desc).toUpperCase().includes(k))
          ],
        });
      } else if (where === null) {
        flags.push({
          desc: clean(t.desc),
          category: t.category,
          amount: t.amount,
          note: (t.place ? `Location "${t.place}" is ` : 'Location ')
              + 'not in the GTA list or the outside-GTA list in rules.json. '
              + 'Left as is - add the city to rules.json to settle it for good.',
        });
      }
    }
  }

  return { transactions, review, flags };
}

/** Fold decisions from a reviewed workbook back into the rules so they stick. */
export function learn(rules, cardId, decisions) {
  rules.merchants[cardId] = rules.merchants[cardId] || {};
  let added = 0;
  for (const { desc, category } of decisions) {
    const key = merchantKey(desc);
    if (!key || !category) continue;
    if (rules.merchants[cardId][key] !== category) {
      rules.merchants[cardId][key] = category;
      added++;
    }
    // a merchant you've now ruled on should stop being force-reviewed
    rules.alwaysReview = (rules.alwaysReview || []).filter(
      (r) => !(key.toUpperCase().includes(r.match.toUpperCase()) && (!r.card || r.card === cardId))
    );
  }
  return added;
}

export function summarize(transactions) {
  const totals = {};
  for (const t of transactions) totals[t.category] = (totals[t.category] || 0) + t.amount;
  const rounded = {};
  for (const k of Object.keys(totals).sort()) rounded[k] = Math.round(totals[k] * 100) / 100;
  return rounded;
}
