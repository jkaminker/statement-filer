import { loadRules, saveRules, resetRules, learn, summarize } from './rules.js';
import { run as runPipeline, inspect as inspectFiles, applyReview } from './pipeline.js';
import { readReviewedWorkbook, readFiledWorkbook } from './workbook.js';
import * as drive from './drive.js';

const $ = (id) => document.getElementById(id);
const LS_CLIENT = 'statement-filer.clientId';
const LS_LAST = 'statement-filer.lastRun';

let rules = null;
let files = [];
let lastResult = null;

// The preflight read of whatever is currently in the file list. Cached against
// the file names and sizes so changing the target quarter, or flipping between
// append and rebuild, never re-reads the PDFs — only a change to the files does.
let inspected = null;
let inspectedKey = '';
let preflightSeq = 0;

// pdf.js is loaded as a module so its worker can be wired up before first use
const pdfjsLib = await import('../vendor/pdf.min.mjs');
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('../vendor/pdf.worker.min.mjs', import.meta.url).href;

const libs = () => ({ pdfjsLib, ExcelJS: window.ExcelJS, PDFLib: window.PDFLib });

const money = (n) =>
  (n < 0 ? '-' : '') + '$' + Math.abs(n).toLocaleString('en-CA', {
    minimumFractionDigits: 2, maximumFractionDigits: 2,
  });

// ══════════════════════════════════════════════════════════════ boot ══
rules = await loadRules();
initTabs();
initSettings();
initDrop();
initPreflight();
initRules();
initReview();
initFiled();
refreshDriveState();

// ══════════════════════════════════════════════════════════════ tabs ══
function initTabs() {
  document.querySelectorAll('.tab').forEach((tab) => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach((t) => t.classList.remove('is-active'));
      document.querySelectorAll('.panel').forEach((p) => p.classList.remove('is-active'));
      tab.classList.add('is-active');
      $(`tab-${tab.dataset.tab}`).classList.add('is-active');
    });
  });
}

// ══════════════════════════════════════════════════════════ settings ══
function initSettings() {
  const dlg = $('settings');
  $('settingsBtn').addEventListener('click', () => {
    $('clientId').value = localStorage.getItem(LS_CLIENT) || '';
    $('fyMonth').value = String(rules.fiscalYearEndMonth || 9);
    dlg.showModal();
  });
  dlg.addEventListener('close', () => {
    if (dlg.returnValue !== 'save') return;
    localStorage.setItem(LS_CLIENT, $('clientId').value.trim());
    rules.fiscalYearEndMonth = Number($('fyMonth').value);
    saveRules(rules);
    refreshDriveState();
  });

  $('connectBtn').addEventListener('click', async () => {
    const clientId = localStorage.getItem(LS_CLIENT);
    if (!drive.isConfigured(clientId)) {
      alert('Add your Google OAuth Client ID first — the ⚙ button, top right. '
          + 'The Setup tab walks through making one.');
      return;
    }
    try {
      await drive.connect(clientId);
      refreshDriveState();
    } catch (e) {
      alert(`Could not connect to Google Drive.\n\n${e.message}`);
    }
  });
}

function refreshDriveState() {
  const on = drive.isSignedIn();
  $('driveState').textContent = on ? 'Drive connected' : 'Drive not connected';
  $('driveState').className = `pill ${on ? 'pill-on' : 'pill-off'}`;
  $('connectBtn').textContent = on ? 'Reconnect' : 'Connect Google Drive';
  $('filedHint').textContent = on
    ? 'Pick a card and quarter, then Open.'
    : 'Connect Google Drive to browse what\'s filed.';
  $('filedLoadBtn').disabled = !on;
  // connecting mid-session is the common case — the lookup can only run now
  if (on && files.length) refreshPreflight();
}

// ═══════════════════════════════════════════════════════ drop & run ══
function initDrop() {
  const drop = $('drop');
  const input = $('fileInput');

  $('pickBtn').addEventListener('click', () => input.click());
  input.addEventListener('change', () => addFiles([...input.files]));

  ['dragenter', 'dragover'].forEach((e) =>
    drop.addEventListener(e, (ev) => { ev.preventDefault(); drop.classList.add('is-over'); }));
  ['dragleave', 'drop'].forEach((e) =>
    drop.addEventListener(e, (ev) => { ev.preventDefault(); drop.classList.remove('is-over'); }));
  drop.addEventListener('drop', (ev) => addFiles([...ev.dataTransfer.files]));

  $('runBtn').addEventListener('click', doRun);
  $('clearBtn').addEventListener('click', () => {
    files = [];
    lastResult = null;
    inspected = null;
    inspectedKey = '';
    renderFiles();
    $('preflight').hidden = true;
    $('result').hidden = true;
    $('log').hidden = true;
  });
}

function addFiles(incoming) {
  for (const f of incoming) {
    if (!/\.pdf$/i.test(f.name)) continue;
    if (files.some((x) => x.name === f.name && x.size === f.size)) continue;
    files.push(f);
  }
  renderFiles();
  refreshPreflight();
}

function renderFiles() {
  const list = $('fileList');
  list.innerHTML = '';
  files.forEach((f, i) => {
    const li = document.createElement('li');
    const name = document.createElement('span');
    name.className = 'name';
    name.textContent = f.name;
    const meta = document.createElement('span');
    meta.className = 'meta';
    meta.textContent = `${(f.size / 1024).toFixed(0)} KB`;
    const x = document.createElement('button');
    x.textContent = '×';
    x.title = 'Remove';
    x.addEventListener('click', () => { files.splice(i, 1); renderFiles(); refreshPreflight(); });
    li.append(name, meta, x);
    list.append(li);
  });
  $('runBtn').disabled = files.length === 0;
  $('clearBtn').disabled = files.length === 0;
}

function logTo(el, msg, isError = false) {
  el.hidden = false;
  const line = document.createElement('div');
  if (isError) line.className = 'err';
  line.textContent = msg;
  el.append(line);
  el.scrollTop = el.scrollHeight;
}

// ═══════════════════════════════════════════════════════════ preflight ══
// Everything the app can tell you before it does any work: which card it sees,
// which quarter these statements belong to, and — the part that used to be
// invisible — whether that quarter already has a workbook in Drive that this run
// is about to add to.

// Whether YOU chose rebuild, as opposed to the app falling back to it because
// there was nothing to append to. Without this, switching the quarter to one
// that IS filed would leave rebuild selected — and quietly overwrite it.
let rebuildChosenByUser = false;

function initPreflight() {
  $('pfQuarter').addEventListener('change', () => {
    // only the Drive side depends on the quarter; the PDFs are already read
    renderQuarterSplit();
    renderPreflightFiled();
  });
  document.querySelectorAll('input[name="runMode"]').forEach((r) => {
    r.addEventListener('change', (ev) => {
      if (ev.isTrusted) rebuildChosenByUser = ev.target.value === 'rebuild';
      updateRunButton();
    });
  });
}

const fileKey = () => files.map((f) => `${f.name}:${f.size}`).join('|');
const chosenQuarter = () => $('pfQuarter').value || (inspected && inspected.quarter) || '';
const chosenMode = () =>
  (document.querySelector('input[name="runMode"]:checked') || {}).value || 'append';

async function refreshPreflight() {
  const panel = $('preflight');
  if (!files.length) {
    panel.hidden = true;
    inspected = null;
    inspectedKey = '';
    updateRunButton();
    return;
  }

  const key = fileKey();
  if (key === inspectedKey && inspected) {
    // same files, nothing to re-read — just refresh the Drive side
    await renderPreflightFiled();
    return;
  }

  const seq = ++preflightSeq;
  panel.hidden = false;
  $('pfBody').hidden = true;
  $('pfBadge').hidden = true;
  $('pfTitle').textContent = files.length === 1
    ? 'Reading the statement…'
    : `Reading ${files.length} statements…`;
  updateRunButton();

  try {
    const seen = await inspectFiles(files, rules, libs(), () => {});
    if (seq !== preflightSeq) return;   // a newer drop overtook this one
    inspected = seen;
    inspectedKey = key;
  } catch (e) {
    if (seq !== preflightSeq) return;
    inspected = null;
    inspectedKey = '';
    $('pfTitle').textContent = 'Could not read those statements';
    $('pfBadge').hidden = false;
    $('pfBadge').className = 'pill pill-bad';
    $('pfBadge').textContent = 'Not parsed';
    $('pfBody').hidden = false;
    $('pfStatements').innerHTML = '';
    $('pfFiled').innerHTML = `<div class="pf-note pf-note-warn">${escapeHtml(e.message)}</div>`;
    $('pfMode').hidden = true;
    $('pfQuarter').innerHTML = '';
    $('pfPath').textContent = '—';
    updateRunButton();
    return;
  }

  $('pfMode').hidden = false;
  $('pfTitle').textContent = `${inspected.cardLabel} · ${inspected.statements.length} statement`
    + `${inspected.statements.length === 1 ? '' : 's'}`;
  $('pfStatements').innerHTML = inspected.statements
    .map((s) => `<span class="chip chip-static">${escapeHtml(s.label)} · ${s.count} txns</span>`)
    .join('');

  await fillQuarterOptions();
  $('pfBody').hidden = false;
  await renderPreflightFiled();
}

/**
 * The quarter picker holds the quarter these statements point at, every quarter
 * this card already has a folder for, and the quarters on either side of the
 * detected one — enough to refile into a neighbouring quarter without offering
 * a list of every quarter that has ever existed.
 */
async function fillQuarterOptions() {
  const sel = $('pfQuarter');
  const wanted = new Set([inspected.quarter, ...inspected.quarters]);
  for (const q of neighbours(inspected.quarter)) wanted.add(q);
  if (drive.isSignedIn()) {
    try {
      for (const q of await drive.listQuarters(inspected.card, inspected.fiscalYear, rules)) {
        wanted.add(q);
      }
    } catch (_) { /* the picker is still usable without Drive */ }
  }
  const list = [...wanted].sort((a, b) => rank(b) - rank(a));
  sel.innerHTML = list
    .map((q) => `<option value="${q}"${q === inspected.quarter ? ' selected' : ''}>${q}</option>`)
    .join('');
  sel.value = inspected.quarter;
  $('pfQuarterNote').textContent = inspected.quarters.length > 1
    ? `these statements span ${inspected.quarters.join(' and ')}`
    : '';
  renderQuarterSplit();
}

/**
 * Say plainly, before Analyze runs, how many rows fall outside the target
 * quarter. The first statement of a quarter always carries some of the previous
 * one, and seeing the number here is what stops it being a surprise later.
 */
function renderQuarterSplit() {
  const host = $('pfSplit');
  if (!host || !inspected) return;
  const q = chosenQuarter();
  const by = inspected.byQuarter || {};
  const inQ = by[q] || 0;
  const out = Object.entries(by).filter(([k]) => k !== q);
  const outCount = out.reduce((s, [, n]) => s + n, 0);

  if (!outCount) {
    host.innerHTML = `<div class="pf-note">All ${inQ} transaction`
      + `${inQ === 1 ? '' : 's'} fall inside ${escapeHtml(q)}.</div>`;
    return;
  }
  host.innerHTML = `<div class="pf-note pf-note-warn"><strong>${inQ} transaction`
    + `${inQ === 1 ? '' : 's'} counted in ${escapeHtml(q)}.</strong> `
    + `${outCount} fall outside it (`
    + out.map(([k, n]) => `${n} in ${escapeHtml(k)}`).join(', ')
    + ') — these go on a separate sheet and are left out of every category total, '
    + 'but they still count toward the reconciliation.</div>';
}

function rank(q) {
  const m = String(q).match(/^Q([1-4]) (\d{4})$/);
  return m ? Number(m[2]) * 10 + Number(m[1]) : 0;
}

function neighbours(q) {
  const m = String(q).match(/^Q([1-4]) (\d{4})$/);
  if (!m) return [];
  let n = Number(m[1]);
  let y = Number(m[2]);
  const out = [];
  for (const step of [-1, 1]) {
    let qq = n + step;
    let yy = y;
    if (qq === 0) { qq = 4; yy--; }
    if (qq === 5) { qq = 1; yy++; }
    out.push(`Q${qq} ${yy}`);
  }
  return out;
}

/** The Drive half of the panel: what this card+quarter already holds. */
async function renderPreflightFiled() {
  const host = $('pfFiled');
  const badge = $('pfBadge');
  if (!inspected) return;
  const quarter = chosenQuarter();

  if (!drive.isSignedIn()) {
    badge.hidden = false;
    badge.className = 'pill pill-warn';
    badge.textContent = 'Drive not connected';
    host.innerHTML = '<div class="pf-note pf-note-warn">'
      + '<strong>Connect Google Drive to add to an existing quarter.</strong> '
      + 'Without it the app can\'t see what\'s already filed, so this run would build '
      + `${escapeHtml(quarter)} from these statements alone.</div>`;
    $('pfPath').textContent = '—';
    setAppendAvailable(false, 'Needs Drive connected.');
    updateRunButton();
    return;
  }

  host.innerHTML = '<div class="pf-note">Checking Drive…</div>';
  const seq = preflightSeq;
  let probe = null;
  try {
    probe = await drive.probeFiled(inspected.card, quarter, inspected.fiscalYear, rules);
  } catch (e) {
    if (seq !== preflightSeq) return;
    host.innerHTML = `<div class="pf-note pf-note-warn">Could not reach Drive: `
      + `${escapeHtml(e.message)}</div>`;
    setAppendAvailable(false, 'Drive lookup failed.');
    updateRunButton();
    return;
  }
  if (seq !== preflightSeq) return;

  $('pfPath').textContent = probe
    ? probe.path
    : `${rules.driveRoot.map((n) => n.replace('{fy}', inspected.fiscalYear)).join(' / ')}`
      + ` / ${rules.cards[inspected.card].driveFolder} / ${quarter}  (will be created)`;

  if (!probe || !probe.workbook) {
    badge.hidden = false;
    badge.className = 'pill pill-on';
    badge.textContent = 'New quarter';
    host.innerHTML = `<div class="pf-note">Nothing filed under <strong>${escapeHtml(quarter)}</strong> `
      + 'yet — this run creates it.</div>';
    setAppendAvailable(false, 'Nothing filed yet, so there is nothing to add to.');
    updateRunButton();
    return;
  }

  // read the filed workbook so the panel can say what is actually in it
  let filed = null;
  try {
    const bytes = await drive.downloadFile(probe.workbook.id);
    if (seq !== preflightSeq) return;
    filed = await readFiledWorkbook(window.ExcelJS, bytes, inspected.card, rules);
  } catch (e) {
    host.innerHTML = `<div class="pf-note pf-note-warn">`
      + `<strong>${escapeHtml(probe.workbook.name)}</strong> is filed, but could not be read: `
      + `${escapeHtml(e.message)}</div>`;
    setAppendAvailable(true, '');
    updateRunButton();
    return;
  }

  const reviewCat = rules.reviewCategory || 'Review';
  const pending = filed.rows.filter((t) => t.category === reviewCat).length;
  const already = new Set(filed.statements.map((s) => s.label));
  const dupes = inspected.statements.filter((s) => already.has(s.label));
  const adding = inspected.statements.filter((s) => !already.has(s.label));

  badge.hidden = false;
  badge.className = 'pill pill-on';
  badge.textContent = 'Already filed';

  const bits = [];
  bits.push(
    `<div class="pf-note"><strong><a href="${probe.workbook.url}" target="_blank" `
    + `rel="noopener">${escapeHtml(probe.workbook.name)}</a></strong> is already filed — `
    + `${filed.rows.length} row${filed.rows.length === 1 ? '' : 's'} covering `
    + `${filed.statements.map((s) => escapeHtml(s.label)).join(' & ') || 'no listed statements'}`
    + (pending
      ? `, <strong>${pending} still awaiting a decision</strong> on the Review sheet`
      : ', nothing outstanding on the Review sheet')
    + '.</div>'
  );
  if (dupes.length) {
    bits.push(
      `<div class="pf-note pf-note-warn">${dupes.map((s) => escapeHtml(s.label)).join(' & ')} `
      + `${dupes.length === 1 ? 'is' : 'are'} already in that workbook and will be skipped, `
      + 'so re-dropping costs nothing.</div>'
    );
  }
  host.innerHTML = bits.join('');

  $('pfAppendHint').textContent = adding.length
    ? `Keeps the ${filed.rows.length} rows already filed and adds `
      + `${adding.map((s) => s.label).join(' & ')}.`
    : 'Everything you\'ve loaded is already filed — this would rebuild the same workbook.';
  setAppendAvailable(true, '');
  updateRunButton();
}

function setAppendAvailable(ok, why) {
  const append = document.querySelector('input[name="runMode"][value="append"]');
  const rebuild = document.querySelector('input[name="runMode"][value="rebuild"]');
  append.disabled = !ok;
  append.closest('.pf-radio').classList.toggle('is-disabled', !ok);
  if (!ok) {
    rebuild.checked = true;
    $('pfAppendHint').textContent = why;
  } else if (!rebuildChosenByUser) {
    // append became possible again (you changed quarter, or connected Drive):
    // fall back to the safe mode unless you asked for rebuild yourself
    append.checked = true;
  }
}

function updateRunButton() {
  const btn = $('runBtn');
  btn.disabled = files.length === 0;
  if (!files.length || !inspected) {
    btn.textContent = 'Analyze';
    return;
  }
  btn.textContent = chosenMode() === 'rebuild'
    ? `Rebuild ${chosenQuarter()}`
    : `Add to ${chosenQuarter()}`;
}

// ═════════════════════════════════════════════════════════════════ run ══
async function doRun() {
  const log = $('log');
  log.innerHTML = '';
  $('result').hidden = true;
  $('runBtn').disabled = true;
  const mode = chosenMode();
  const quarter = chosenQuarter();
  try {
    lastResult = await runPipeline(files, rules, libs(), (m) => logTo(log, m), {
      // merge by default: if this card and quarter are already filed, add to it
      fetchFiled: drive.isSignedIn()
        ? (card, q, fy) => drive.fetchFiled(card, q, fy, rules, (m) => logTo(log, m))
        : null,
      inspected: inspectedKey === fileKey() ? inspected : null,
      quarter,
      mode,
    });
    logTo(log, 'Done.');
    renderResult(lastResult, $('result'), false);
    publishForTests(lastResult);
  } catch (e) {
    logTo(log, e.message, true);
    console.error(e);
  } finally {
    updateRunButton();
  }
}

// ═══════════════════════════════════════════════════════════ results ══
function renderResult(res, host, isReview) {
  host.hidden = false;
  host.innerHTML = '';

  // reconciliation banner — the single most important thing on the page
  const tied = res.controlTotal && Math.abs(res.variance) < 0.005;
  const banner = document.createElement('div');
  const outsideBit = res.outside && res.outside.length
    ? ` Plus ${res.outside.length} transaction${res.outside.length === 1 ? '' : 's'} `
      + `totalling ${money(res.outsideTotal)} that fall outside ${res.quarter} and are `
      + 'listed separately.'
    : '';
  if (tied) {
    banner.className = 'banner banner-good';
    banner.innerHTML = `<span>✓</span><div><strong>Reconciled.</strong> `
      + `${res.transactions.length} transactions totalling ${money(res.parsedTotal)} `
      + `in ${escapeHtml(res.quarter)}.${outsideBit} `
      + 'Together these match the statement totals exactly.</div>';
  } else if (res.controlTotal) {
    // Show every term of the sum. The quarter total alone does not equal the
    // statement total whenever rows fall outside the quarter, so printing just
    // those two makes the difference look like arithmetic that doesn't work.
    banner.className = 'banner banner-bad';
    const sum = res.outside && res.outside.length
      ? `${money(res.parsedTotal)} in ${escapeHtml(res.quarter)} plus `
        + `${money(res.outsideTotal)} outside it = ${money(res.parsedTotal + res.outsideTotal)}`
      : `${money(res.parsedTotal)}`;
    banner.innerHTML = `<span>!</span><div><strong>Does not reconcile.</strong> `
      + `${sum}, against a statement total of ${money(res.controlTotal)} — `
      + `a difference of ${money(res.variance)}. Don't file this until we work out why.</div>`;
  } else {
    banner.className = 'banner banner-warn';
    banner.innerHTML = `<span>?</span><div><strong>No control total found</strong> on these `
      + `statements, so I can't verify the parse. Parsed ${money(res.parsedTotal)}.</div>`;
  }
  host.append(banner);

  const card = document.createElement('div');
  card.className = 'card';
  card.innerHTML = `<h2>${res.cardLabel} · ${res.quarter}</h2>`
    + `<p class="muted small">Filed under fiscal year ending September ${res.fiscalYear}`
    + (res.quarterOverridden ? ` · quarter chosen by hand (auto would be ${res.quarterAuto})` : '')
    + '.</p>'
    + mergeLine(res);

  const wrap = document.createElement('div');
  wrap.className = 'table-wrap';
  const rows = Object.entries(res.summary)
    .map(([k, v]) => `<tr><td>${escapeHtml(k)}</td><td class="num">${money(v)}</td></tr>`)
    .join('');
  wrap.innerHTML = `<table><thead><tr><th>Category</th><th class="num">Amount</th></tr></thead>`
    + `<tbody>${rows}<tr class="total"><td>Grand Total</td>`
    + `<td class="num">${money(res.parsedTotal)}</td></tr></tbody></table>`;
  card.append(wrap);

  if (isReview && res.applied !== undefined) {
    const p = document.createElement('p');
    p.className = 'muted small';
    p.style.marginTop = '.8rem';
    p.textContent = `${res.applied} item${res.applied === 1 ? '' : 's'} recoded from your comments.`
      + (res.moved.length ? ` ${res.moved.length} already-categorized item(s) also moved.` : '');
    card.append(p);
  }

  if (res.review.length) {
    const p = document.createElement('p');
    p.className = 'muted small';
    p.style.marginTop = '.8rem';
    p.innerHTML = `<strong>${res.review.length} item${res.review.length === 1 ? '' : 's'} need`
      + `${res.review.length === 1 ? 's' : ''} your decision.</strong> They're on the Review sheet, `
      + `coded <code>Review</code> and kept out of every other category. Fill in the COMMENTS `
      + `column and come back to tab 2.`;
    card.append(p);
  }

  // downloads
  const chips = document.createElement('div');
  chips.className = 'chips';
  chips.append(downloadChip(res.workbook.name, res.workbook.bytes,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'));
  for (const c of res.categoryPdfs) {
    chips.append(downloadChip(`${c.category} (${c.lines})`, c.bytes, 'application/pdf', c.fileName));
  }
  const label = document.createElement('p');
  label.className = 'muted small';
  label.style.margin = '1rem 0 0';
  label.textContent = 'Download individually:';
  card.append(label, chips);

  const actions = document.createElement('div');
  actions.className = 'actions';
  const fileBtn = document.createElement('button');
  fileBtn.className = 'btn btn-primary';
  fileBtn.textContent = 'File everything to Google Drive';
  fileBtn.disabled = !tied;
  fileBtn.title = tied ? '' : 'Reconcile first';
  fileBtn.addEventListener('click', () => doFile(res, host));
  actions.append(fileBtn);
  card.append(actions);

  host.append(card);
}

/** One line saying, in plain terms, what this run did to the quarter. */
function mergeLine(res) {
  const m = res.merged || {};
  if (res.mode === 'rebuild') {
    return '<p class="pf-note pf-note-warn">Rebuilt from scratch — '
      + `${res.transactions.length} rows from the statements loaded. Anything previously filed `
      + `under ${escapeHtml(res.quarter)} has been replaced.</p>`;
  }
  if (!m.carried) return '';
  const parts = [`<strong>${m.carried}</strong> row${m.carried === 1 ? '' : 's'} carried forward`,
                 `<strong>${m.added}</strong> added this run`];
  if (m.skipped && m.skipped.length) {
    parts.push(`${m.skipped.map(escapeHtml).join(' & ')} skipped as already filed`);
  }
  if (m.carriedReview) {
    parts.push(`<strong>${m.carriedReview}</strong> older item`
      + `${m.carriedReview === 1 ? '' : 's'} still awaiting your decision`);
  }
  return `<p class="pf-note">Added to the quarter: ${parts.join(', ')}.</p>`;
}

function downloadChip(text, bytes, mime, fileName) {
  const a = document.createElement('a');
  a.className = 'chip';
  a.href = URL.createObjectURL(new Blob([bytes], { type: mime }));
  a.download = fileName || text;
  a.textContent = `↓ ${text}`;
  return a;
}

async function doFile(res, host) {
  if (!drive.isSignedIn()) {
    alert('Connect Google Drive first — the button at the top right.');
    return;
  }
  const log = $(host.id === 'result' ? 'log' : 'reviewLog');
  try {
    const out = await drive.fileRun(res, rules, (m) => logTo(log, m));
    localStorage.setItem(LS_LAST, JSON.stringify({
      card: res.card, quarter: res.quarter, folderId: out.folderId,
      workbookName: res.workbook.name,
    }));
    const done = document.createElement('div');
    done.className = 'banner banner-good';
    done.innerHTML = `<span>✓</span><div><strong>Filed.</strong> `
      + `${out.uploaded.length} files are in <code>${escapeHtml(out.path)}</code>. `
      + `<a href="${out.folderUrl}" target="_blank" rel="noopener">Open the folder in Drive</a>.</div>`;
    host.prepend(done);
  } catch (e) {
    logTo(log, e.message, true);
    alert(`Filing to Drive failed.\n\n${e.message}`);
  }
}

// ════════════════════════════════════════════════════════════ review ══
function initReview() {
  $('reviewPickBtn').addEventListener('click', () => $('reviewInput').click());
  $('reviewInput').addEventListener('change', async () => {
    const f = $('reviewInput').files[0];
    if (f) doApplyReview(await f.arrayBuffer());
  });
  $('reviewDriveBtn').addEventListener('click', async () => {
    const last = JSON.parse(localStorage.getItem(LS_LAST) || 'null');
    if (!last) { alert('No previous run recorded in this browser. Choose the file instead.'); return; }
    if (!drive.isSignedIn()) { alert('Connect Google Drive first.'); return; }
    try {
      const file = await drive.findFile(last.workbookName, last.folderId);
      if (!file) { alert(`Could not find ${last.workbookName} in Drive.`); return; }
      doApplyReview(await drive.downloadFile(file.id));
    } catch (e) {
      alert(e.message);
    }
  });
}

async function doApplyReview(arrayBuffer) {
  const log = $('reviewLog');
  log.innerHTML = '';
  $('reviewResult').hidden = true;
  if (!lastResult) {
    logTo(log, 'Load the statement PDFs on tab 1 and run Analyze first — '
      + 'the highlighted files have to be redrawn from the original statements.', true);
    return;
  }
  try {
    logTo(log, 'Reading your comments…');
    const decisions = await readReviewedWorkbook(
      window.ExcelJS, arrayBuffer, lastResult.card, rules
    );
    logTo(log, `${decisions.rowDecisions.length} row decision(s), `
      + `${decisions.merchantDecisions.length} merchant confirmation(s).`);

    const updated = await applyReview(lastResult, decisions, rules, libs(), (m) => logTo(log, m));

    if ($('learnCheck').checked) {
      const all = [...decisions.rowDecisions, ...decisions.merchantDecisions];
      const added = learn(rules, lastResult.card, all);
      saveRules(rules);
      initRules();
      if (added) logTo(log, `${added} merchant rule(s) remembered for next time.`);
    }

    lastResult = updated;
    logTo(log, 'Done.');
    renderResult(updated, $('reviewResult'), true);
    publishForTests(updated);
  } catch (e) {
    logTo(log, e.message, true);
    console.error(e);
  }
}

// ════════════════════════════════════════════════════════ what's filed ══
// A read-only window onto a quarter. Nothing here writes to Drive — it exists
// so you can answer "what's in Q3 and what's still waiting on me?" without
// loading a single PDF.

function initFiled() {
  const cardSel = $('filedCard');
  cardSel.innerHTML = Object.entries(rules.cards)
    .map(([id, c]) => `<option value="${id}">${escapeHtml(c.label)}</option>`).join('');

  const thisYear = new Date().getFullYear();
  $('filedFy').innerHTML = [thisYear + 1, thisYear, thisYear - 1, thisYear - 2]
    .map((y) => `<option value="${y}"${y === thisYear ? ' selected' : ''}>`
      + `${rules.fiscalFolderTemplate.replace('{fy}', y)}</option>`).join('');

  const reload = () => fillFiledQuarters();
  cardSel.addEventListener('change', reload);
  $('filedFy').addEventListener('change', reload);
  $('filedLoadBtn').addEventListener('click', doFiledLoad);

  document.querySelector('.tab[data-tab="filed"]').addEventListener('click', () => {
    if (drive.isSignedIn() && !$('filedQuarter').options.length) fillFiledQuarters();
  });
}

async function fillFiledQuarters() {
  const sel = $('filedQuarter');
  if (!drive.isSignedIn()) { sel.innerHTML = ''; return; }
  sel.innerHTML = '<option>…</option>';
  try {
    const qs = await drive.listQuarters($('filedCard').value, Number($('filedFy').value), rules);
    sel.innerHTML = qs.length
      ? qs.map((q) => `<option value="${q}">${q}</option>`).join('')
      : '<option value="">nothing filed yet</option>';
    $('filedHint').textContent = qs.length
      ? 'Pick a quarter, then Open.'
      : 'No quarters filed under that card and audit year yet.';
  } catch (e) {
    sel.innerHTML = '';
    $('filedHint').textContent = `Could not list quarters: ${e.message}`;
  }
}

async function doFiledLoad() {
  const log = $('filedLog');
  const host = $('filedResult');
  log.innerHTML = '';
  host.hidden = true;
  const card = $('filedCard').value;
  const fy = Number($('filedFy').value);
  const quarter = $('filedQuarter').value;
  if (!quarter) { logTo(log, 'Pick a quarter first.', true); return; }

  $('filedLoadBtn').disabled = true;
  try {
    logTo(log, `Looking up ${rules.cards[card].label} ${quarter}…`);
    const probe = await drive.probeFiled(card, quarter, fy, rules);
    if (!probe || !probe.workbook) {
      logTo(log, `No workbook filed under ${quarter} yet.`, true);
      return;
    }
    logTo(log, `Reading ${probe.workbook.name}…`);
    const bytes = await drive.downloadFile(probe.workbook.id);
    const filed = await readFiledWorkbook(window.ExcelJS, bytes, card, rules);
    logTo(log, 'Done.');
    renderFiled(host, { card, quarter, probe, filed });
  } catch (e) {
    logTo(log, e.message, true);
    console.error(e);
  } finally {
    $('filedLoadBtn').disabled = !drive.isSignedIn();
  }
}

function renderFiled(host, { card, quarter, probe, filed }) {
  host.hidden = false;
  host.innerHTML = '';
  const reviewCat = rules.reviewCategory || 'Review';
  const pending = filed.rows.filter((t) => t.category === reviewCat);
  const total = filed.rows.reduce((s, t) => s + t.amount, 0);

  const banner = document.createElement('div');
  banner.className = pending.length ? 'banner banner-warn' : 'banner banner-good';
  banner.innerHTML = pending.length
    ? `<span>!</span><div><strong>${pending.length} item`
      + `${pending.length === 1 ? '' : 's'} waiting on you.</strong> `
      + `Open the workbook, fill in the COMMENTS column on the Review sheet, then bring it `
      + 'back to the "Apply my review" tab.</div>'
    : `<span>✓</span><div><strong>Nothing outstanding.</strong> Every row in this quarter `
      + 'has a category.</div>';
  host.append(banner);

  const c = document.createElement('div');
  c.className = 'card';
  c.innerHTML = `<h2>${escapeHtml(rules.cards[card].label)} · ${escapeHtml(quarter)}</h2>`
    + `<p class="muted small mono">${escapeHtml(probe.path)}</p>`
    + `<p class="muted">${filed.rows.length} rows totalling ${money(Math.round(total * 100) / 100)}, `
    + `covering ${filed.statements.map((s) => escapeHtml(s.label)).join(' & ') || '—'}.</p>`;

  const wrap = document.createElement('div');
  wrap.className = 'table-wrap';
  const summary = summarize(filed.rows);
  wrap.innerHTML = '<table><thead><tr><th>Category</th><th class="num">Amount</th></tr></thead>'
    + '<tbody>'
    + Object.entries(summary)
      .map(([k, v]) => `<tr${k === reviewCat ? ' class="is-review"' : ''}>`
        + `<td>${escapeHtml(k)}</td><td class="num">${money(v)}</td></tr>`).join('')
    + `<tr class="total"><td>Grand Total</td>`
    + `<td class="num">${money(Math.round(total * 100) / 100)}</td></tr></tbody></table>`;
  c.append(wrap);

  if (pending.length) {
    const pw = document.createElement('div');
    pw.className = 'table-wrap';
    pw.innerHTML = '<h3>Still in Review</h3>'
      + '<table><thead><tr><th>Date</th><th>Merchant</th><th class="num">Amount</th>'
      + '<th>Note</th></tr></thead><tbody>'
      + pending.map((t) => `<tr><td>${escapeHtml(t.date)}</td>`
        + `<td>${escapeHtml(t.desc)}</td><td class="num">${money(t.amount)}</td>`
        + `<td class="small muted">${escapeHtml(t.note || '')}</td></tr>`).join('')
      + '</tbody></table>';
    c.append(pw);
  }

  const actions = document.createElement('div');
  actions.className = 'actions';
  const openWb = document.createElement('a');
  openWb.className = 'btn';
  openWb.target = '_blank';
  openWb.rel = 'noopener';
  openWb.href = probe.workbook.url;
  openWb.textContent = 'Open the workbook in Drive';
  const openFolder = document.createElement('a');
  openFolder.className = 'btn btn-ghost';
  openFolder.target = '_blank';
  openFolder.rel = 'noopener';
  openFolder.href = probe.folderUrl;
  openFolder.textContent = 'Open the quarter folder';
  actions.append(openWb, openFolder);
  c.append(actions);

  host.append(c);
}

// ═════════════════════════════════════════════════════════════ rules ══
function initRules() {
  $('rulesText').value = JSON.stringify(rules, null, 1);
  const merchants = Object.values(rules.merchants || {})
    .reduce((n, m) => n + Object.keys(m).length, 0);
  $('rulesSummary').innerHTML = `
    <div><strong>${merchants}</strong>merchant rules</div>
    <div><strong>${(rules.alwaysReview || []).length}</strong>always-review entries</div>
    <div><strong>${(rules.gtaRule?.gtaPlaces || []).length}</strong>GTA places</div>
    <div><strong>${(rules.gtaRule?.outsidePlaces || []).length}</strong>outside-GTA places</div>`;
  $('rulesPrivacy').textContent = merchants === 0
    ? 'No merchant rules loaded in this browser yet — import my-rules.json below, or just run a '
      + 'quarter and answer the Review sheet, and it will learn as you go.'
    : '';

  $('rulesSaveBtn').onclick = () => {
    try {
      rules = JSON.parse($('rulesText').value);
      saveRules(rules);
      initRules();
      $('rulesMsg').textContent = 'Saved to this browser. Download and commit it to your repo '
        + 'to make it permanent across devices.';
    } catch (e) {
      $('rulesMsg').textContent = `That isn't valid JSON: ${e.message}`;
    }
  };
  $('rulesImportBtn').onclick = () => $('rulesImportInput').click();
  $('rulesImportInput').onchange = async () => {
    const f = $('rulesImportInput').files[0];
    if (!f) return;
    try {
      const incoming = JSON.parse(await f.text());
      if (!incoming.merchants || !incoming.cards) {
        throw new Error("that doesn't look like a rules file (no merchants or cards section).");
      }
      // merge merchants into whatever is loaded, so importing tops up rather than wipes
      for (const card of Object.keys(incoming.merchants)) {
        rules.merchants[card] = { ...(rules.merchants[card] || {}), ...incoming.merchants[card] };
      }
      for (const key of ['alwaysReview', 'spendCategoryDefaults', 'gtaRule',
                         'largeAmountReview', 'cards', 'fiscalYearEndMonth']) {
        if (incoming[key] !== undefined) rules[key] = incoming[key];
      }
      saveRules(rules);
      initRules();
      const n = Object.values(rules.merchants).reduce((a, m) => a + Object.keys(m).length, 0);
      $('rulesMsg').textContent = `Imported. ${n} merchant rules are now loaded in this browser.`;
    } catch (e) {
      $('rulesMsg').textContent = `Could not import that file: ${e.message}`;
    }
    $('rulesImportInput').value = '';
  };
  $('rulesDownloadBtn').onclick = () => {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([JSON.stringify(rules, null, 1)],
      { type: 'application/json' }));
    a.download = 'rules.json';
    a.click();
  };
  $('rulesResetBtn').onclick = async () => {
    resetRules();
    rules = await loadRules();
    initRules();
    $('rulesMsg').textContent = 'Reset to the copy in your repo.';
  };
}

/** A small, serializable snapshot of a run, so the test harness can assert on it. */
function publishForTests(res) {
  window.__pipelineResult = res;   // full object, for the artifact harness
  window.__lastResult = {
    card: res.card,
    quarter: res.quarter,
    fiscalYear: res.fiscalYear,
    count: res.transactions.length,
    total: res.parsedTotal,
    control: res.controlTotal,
    variance: res.variance,
    categories: res.summary,
    review: res.review.length,
    flags: res.flags.length,
    pdfs: res.categoryPdfs.map((p) => `${p.category}:${p.lines}`),
    workbookBytes: res.workbook.bytes.length,
  };
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"]/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
