import * as amex from './amex.js';
import * as cibc from './cibc.js';

export const PARSERS = [amex, cibc];

export function parserFor(cardId) {
  return PARSERS.find((p) => p.id === cardId) || null;
}

/** Identify which card a statement belongs to. Returns null if nothing matches. */
export function detectCard(pages) {
  for (const p of PARSERS) {
    try {
      if (p.detect(pages)) return p;
    } catch (_) { /* a parser that throws on detect just doesn't match */ }
  }
  return null;
}
