import { loadRules, saveRules, resetRules, learn } from './rules.js';
import { run as runPipeline, applyReview } from './pipeline.js';
import { readReviewedWorkbook } from './workbook.js';
import * as drive from './drive.js';

const $ = (id) => document.getElementById(id);
const LS_CLIENT = 'statement-filer.clientId';
const LS_LAST = 'statement-filer.lastRun';

let rules = null;
let files = [];
let lastResult = null;

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
initRules();
initReview();
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
    renderFiles();
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
    x.addEventListener('click', () => { files.splice(i, 1); renderFiles(); });
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

async function doRun() {
  const log = $('log');
  log.innerHTML = '';
  $('result').hidden = true;
  $('runBtn').disabled = true;
  try {
    lastResult = await runPipeline(files, rules, libs(), (m) => logTo(log, m), {
      // merge by default: if this card and quarter are already filed, add to it
      fetchFiled: drive.isSignedIn()
        ? (card, quarter, fy) => drive.fetchFiled(card, quarter, fy, rules, (m) => logTo(log, m))
        : null,
    });
    logTo(log, 'Done.');
    renderResult(lastResult, $('result'), false);
    publishForTests(lastResult);
  } catch (e) {
    logTo(log, e.message, true);
    console.error(e);
  } finally {
    $('runBtn').disabled = files.length === 0;
  }
}

// ═══════════════════════════════════════════════════════════ results ══
function renderResult(res, host, isReview) {
  host.hidden = false;
  host.innerHTML = '';

  // reconciliation banner — the single most important thing on the page
  const tied = res.controlTotal && Math.abs(res.variance) < 0.005;
  const banner = document.createElement('div');
  if (tied) {
    banner.className = 'banner banner-good';
    banner.innerHTML = `<span>✓</span><div><strong>Reconciled.</strong> `
      + `${res.transactions.length} transactions totalling ${money(res.parsedTotal)}, `
      + `matching the statement totals exactly.</div>`;
  } else if (res.controlTotal) {
    banner.className = 'banner banner-bad';
    banner.innerHTML = `<span>!</span><div><strong>Does not reconcile.</strong> `
      + `Parsed ${money(res.parsedTotal)} against a statement total of ${money(res.controlTotal)} — `
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
    + `<p class="muted small">Filed under fiscal year ending September ${res.fiscalYear}.</p>`;

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
