// Posting a finished quarter into Wave.
//
// What this books, and why it looks like this:
//
// The cards are paid through the shareholder loan, so each category becomes one
// money transaction ON that loan account — a withdrawal — with the expense
// account as its single line item. That is the same shape as the Home Office
// Expense entries already in the ledger, and it is deliberately NOT a journal
// entry: Wave's journal-entry picker excludes bank-type accounts, so
// "Shareholder Loan - Jeff Kaminker" is unreachable from there.
//
// Personal spending is never posted. It is not reclassified, not netted, not
// sent anywhere — it simply does not appear, because the company did not incur
// it.
//
// The token lives in this module and nowhere else. It is never written to
// localStorage, never put in a URL, and is dropped when the tab closes. A Wave
// Full Access Token grants complete access to the books, so the cost of
// re-pasting it each session is cheaper than the cost of storing it.

import { quarterBounds } from './parsers/base.js';

const ENDPOINT = 'https://gql.waveapps.com/graphql/public';

let token = null;

export function isConnected() {
  return !!token;
}

export function signOut() {
  token = null;
}

async function gql(query, variables = {}) {
  if (!token) throw new Error('Not connected to Wave.');
  let res;
  try {
    res = await fetch(ENDPOINT, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${token}`,
      },
      body: JSON.stringify({ query, variables }),
    });
  } catch (e) {
    throw new Error(`Could not reach Wave: ${e.message}`);
  }
  if (res.status === 401 || res.status === 403) {
    token = null;
    throw new Error('Wave rejected that token. Paste a fresh Full Access Token and try again.');
  }
  const body = await res.json().catch(() => null);
  if (!body) throw new Error(`Wave returned ${res.status} with a response this app could not read.`);
  if (body.errors && body.errors.length) {
    throw new Error(body.errors.map((e) => e.message).join(' | '));
  }
  return body.data;
}

/** Validates the token by asking who it belongs to. Returns the business list. */
export async function connect(rawToken) {
  const t = String(rawToken || '').trim();
  if (!t) throw new Error('Paste your Wave Full Access Token first.');
  token = t;
  try {
    const d = await gql(`{ businesses(page:1, pageSize:50) { edges { node { id name isArchived } } } }`);
    const list = (d.businesses?.edges || [])
      .map((e) => e.node)
      .filter((b) => !b.isArchived);
    if (!list.length) throw new Error('That token works, but it can\'t see any businesses.');
    return list;
  } catch (e) {
    token = null;
    throw e;
  }
}

/** Every account this business has, as a name -> id map plus the raw list. */
export async function listAccounts(businessId) {
  const d = await gql(
    `query A($id: ID!) {
       business(id: $id) {
         accounts(page: 1, pageSize: 500) {
           edges { node { id name type { value } subtype { value } } }
         }
       }
     }`,
    { id: businessId }
  );
  const nodes = (d.business?.accounts?.edges || []).map((e) => e.node);
  const byName = {};
  for (const a of nodes) byName[a.name.trim()] = a.id;
  return { byName, accounts: nodes };
}

const slug = (s) =>
  String(s).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');

/**
 * Turn a finished run into the exact list of transactions to post.
 *
 * Pure and synchronous on purpose: everything the preview shows, and every
 * test, runs through this one function, so what you see is what gets posted.
 *
 * Returns { entries, excluded, unmapped, total }.
 *  - entries   : ready to post
 *  - excluded  : categories deliberately left out (Personal, Review)
 *  - unmapped  : categories with no Wave account — these BLOCK posting rather
 *                than being dropped, because a silently vanishing category is
 *                how money goes missing.
 */
export function planEntries(result, rules) {
  const cfg = (rules && rules.wave) || {};
  const map = cfg.accounts || {};
  const skip = new Set([...(cfg.exclude || []), rules?.reviewCategory].filter(Boolean));

  const bounds = quarterBounds(result.quarter);
  const date = bounds ? bounds.to : null;
  if (!date) throw new Error(`Could not work out the quarter-end date for "${result.quarter}".`);

  const label = result.cardLabel || result.card;
  const entries = [];
  const excluded = [];
  const unmapped = [];

  for (const [category, raw] of Object.entries(result.summary || {})) {
    const amount = Math.round(raw * 100) / 100;
    if (!amount) continue;

    if (skip.has(category)) {
      excluded.push({ category, amount });
      continue;
    }
    const account = map[category];
    if (!account) {
      unmapped.push({ category, amount });
      continue;
    }

    // A category that nets negative (more credits than charges) reverses:
    // money comes back to the loan account and the expense is relieved.
    const positive = amount > 0;
    entries.push({
      category,
      account,
      amount: Math.abs(amount),
      date,
      description: `${result.quarter} ${label} Expenses - ${category}`,
      externalId: `sf-${slug(result.card)}-${slug(result.quarter)}-${slug(category)}`,
      direction: positive ? 'WITHDRAWAL' : 'DEPOSIT',
      balance: positive ? 'DEBIT' : 'CREDIT',
    });
  }

  entries.sort((a, b) => b.amount - a.amount);
  const total = Math.round(entries.reduce((s, e) => s + (e.direction === 'WITHDRAWAL' ? e.amount : -e.amount), 0) * 100) / 100;
  return { entries, excluded, unmapped, total };
}

/**
 * Post the planned entries. One bulk call, not one call per line.
 *
 * externalId makes this safe to run twice: Wave treats a repeat as the same
 * transaction rather than a second one, so a double click or a re-run after a
 * dropped connection does not duplicate the quarter.
 */
export async function post(businessId, entries, accountsByName, anchorName) {
  const anchorId = accountsByName[anchorName];
  if (!anchorId) {
    throw new Error(`Wave has no account named "${anchorName}". Check the wave.anchorAccount setting in your rules.`);
  }
  const missing = entries.filter((e) => !accountsByName[e.account]);
  if (missing.length) {
    throw new Error(
      'These expense accounts do not exist in Wave: '
      + [...new Set(missing.map((m) => m.account))].join(', ')
      + '. Create them in Wave, or change the mapping in your rules.'
    );
  }

  const transactions = entries.map((e) => ({
    externalId: e.externalId,
    date: e.date,
    description: e.description,
    anchor: { accountId: anchorId, amount: e.amount, direction: e.direction },
    lineItems: [{ accountId: accountsByName[e.account], amount: e.amount, balance: e.balance }],
  }));

  const d = await gql(
    `mutation Post($input: MoneyTransactionsCreateInput!) {
       moneyTransactionsCreate(input: $input) {
         didSucceed
         inputErrors { path message code }
       }
     }`,
    { input: { businessId, transactions } }
  );

  const r = d.moneyTransactionsCreate || {};
  if (!r.didSucceed) {
    const errs = (r.inputErrors || []).map((e) => `${(e.path || []).join('.')}: ${e.message}`);
    throw new Error(errs.length ? errs.join(' | ') : 'Wave rejected the transactions without saying why.');
  }
  return { posted: transactions.length };
}
