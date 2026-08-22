// Google Drive, straight from the browser. Uses Google Identity Services for a
// token (no client secret, nothing stored on a server) and the Drive REST API to
// create folders and upload files.
//
// Scope is drive.file: the app can only see and touch files it created itself,
// plus anything you explicitly pick. It cannot read the rest of your Drive.

const SCOPE = 'https://www.googleapis.com/auth/drive';
const API = 'https://www.googleapis.com/drive/v3';
const UPLOAD = 'https://www.googleapis.com/upload/drive/v3/files';
const FOLDER_MIME = 'application/vnd.google-apps.folder';

let tokenClient = null;
let accessToken = null;
let tokenExpiry = 0;

export function isConfigured(clientId) {
  return !!clientId && clientId !== 'PASTE_YOUR_CLIENT_ID_HERE';
}

export function isSignedIn() {
  return !!accessToken && Date.now() < tokenExpiry;
}

export function signOut() {
  if (accessToken && window.google?.accounts?.oauth2) {
    window.google.accounts.oauth2.revoke(accessToken, () => {});
  }
  accessToken = null;
  tokenExpiry = 0;
}

/** Opens Google's sign-in popup and holds the token in memory only. */
export function connect(clientId) {
  return new Promise((resolve, reject) => {
    if (!window.google?.accounts?.oauth2) {
      reject(new Error('Google sign-in did not load. Check your connection and reload.'));
      return;
    }
    tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: SCOPE,
      callback: (resp) => {
        if (resp.error) {
          reject(new Error(resp.error_description || resp.error));
          return;
        }
        accessToken = resp.access_token;
        tokenExpiry = Date.now() + (Number(resp.expires_in) - 60) * 1000;
        resolve(true);
      },
      error_callback: (err) => reject(new Error(err.message || 'Sign-in was cancelled.')),
    });
    tokenClient.requestAccessToken({ prompt: '' });
  });
}

async function api(path, opts = {}) {
  const res = await fetch(`${API}${path}`, {
    ...opts,
    headers: { Authorization: `Bearer ${accessToken}`, ...(opts.headers || {}) },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Drive said ${res.status}: ${body.slice(0, 300)}`);
  }
  return res.json();
}

const esc = (s) => s.replace(/'/g, "\\'");

/** Find a folder by name under a parent. Returns its id, or null. */
export async function findFolder(name, parentId) {
  const q = encodeURIComponent(
    `name = '${esc(name)}' and mimeType = '${FOLDER_MIME}' and trashed = false`
    + (parentId ? ` and '${esc(parentId)}' in parents` : '')
  );
  const r = await api(`/files?q=${q}&fields=files(id,name)&pageSize=10`);
  return r.files.length ? r.files[0].id : null;
}

export async function findFile(name, parentId) {
  const q = encodeURIComponent(
    `name = '${esc(name)}' and trashed = false and '${esc(parentId)}' in parents`
  );
  const r = await api(`/files?q=${q}&fields=files(id,name,modifiedTime)&pageSize=10`);
  return r.files.length ? r.files[0] : null;
}

export async function ensureFolder(name, parentId) {
  const existing = await findFolder(name, parentId);
  if (existing) return existing;
  const r = await api('/files?fields=id', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      name,
      mimeType: FOLDER_MIME,
      ...(parentId ? { parents: [parentId] } : {}),
    }),
  });
  return r.id;
}

/** Walk (and create as needed) a chain of folder names. Returns the last id. */
export async function ensurePath(names, rootId = null) {
  let parent = rootId;
  const trail = [];
  for (const n of names) {
    parent = await ensureFolder(n, parent);
    trail.push({ name: n, id: parent });
  }
  return { id: parent, trail };
}

/** Create or overwrite a file in a folder. Returns {id, name, webViewLink}. */
export async function uploadFile(name, mimeType, bytes, parentId) {
  const existing = await findFile(name, parentId);
  const metadata = existing ? { name } : { name, parents: [parentId] };
  const boundary = 'sf' + Math.random().toString(36).slice(2);
  const body = new Blob([
    `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n`,
    JSON.stringify(metadata),
    `\r\n--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`,
    bytes,
    `\r\n--${boundary}--`,
  ]);
  const url = existing
    ? `${UPLOAD}/${existing.id}?uploadType=multipart&fields=id,name,webViewLink`
    : `${UPLOAD}?uploadType=multipart&fields=id,name,webViewLink`;
  const res = await fetch(url, {
    method: existing ? 'PATCH' : 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': `multipart/related; boundary=${boundary}`,
    },
    body,
  });
  if (!res.ok) {
    throw new Error(`Upload of ${name} failed (${res.status}): ${(await res.text()).slice(0, 300)}`);
  }
  const out = await res.json();
  return { ...out, replaced: !!existing };
}

export async function downloadFile(fileId) {
  const res = await fetch(`${API}/files/${fileId}?alt=media`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error(`Could not download that file (${res.status}).`);
  return res.arrayBuffer();
}

/**
 * File a completed run into Drive, mirroring how the audit folder is already
 * laid out on disk — the untouched statements live loose in the card folder,
 * and each quarter keeps only the abridged copy:
 *
 *   Annual Audit {fy} Sep / Credit Card Statements / {Card} /
 *      <the statements exactly as the bank issued them>
 *      {Quarter} /
 *         {Quarter} workbook.xlsx
 *         Abridged Statements/  (transaction pages only)
 *         Expense(s) Summary/   (one highlighted PDF per category)
 */
export async function fileRun(result, rules, onProgress = () => {}) {
  const cfg = rules.cards[result.card];
  const rootNames = rules.driveRoot.map((n) => n.replace('{fy}', result.fiscalYear));

  onProgress('Locating the audit folder…');
  const { id: cardsRoot, trail } = await ensurePath(rootNames);
  const cardFolder = await ensureFolder(cfg.driveFolder, cardsRoot);
  const quarterFolder = await ensureFolder(result.quarter, cardFolder);
  const stmtFolder = await ensureFolder(cfg.statementsFolder, quarterFolder);
  const sumFolder = await ensureFolder(cfg.summaryFolder, quarterFolder);

  const uploaded = [];

  onProgress('Uploading the workbook…');
  uploaded.push(await uploadFile(
    result.workbook.name,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    result.workbook.bytes,
    quarterFolder
  ));

  // one untouched copy of each statement, kept in the card folder alongside the
  // other quarters — skipped if it's already there, so re-running is harmless
  for (const s of result.sources) {
    if (await findFile(s.name, cardFolder)) {
      onProgress(`${s.name} is already filed — leaving it alone.`);
      continue;
    }
    onProgress(`Uploading ${s.name}…`);
    uploaded.push(await uploadFile(s.name, 'application/pdf', s.bytes, cardFolder));
  }

  // the quarter keeps the abridged copy — transaction pages only
  const statementFiles = (result.abridged && result.abridged.length)
    ? result.abridged
    : result.sources;
  for (const s of statementFiles) {
    onProgress(`Uploading ${s.name} (abridged)…`);
    uploaded.push(await uploadFile(s.name, 'application/pdf', s.bytes, stmtFolder));
  }

  for (const c of result.categoryPdfs) {
    onProgress(`Uploading ${c.fileName}…`);
    uploaded.push(await uploadFile(c.fileName, 'application/pdf', c.bytes, sumFolder));
  }

  return {
    uploaded,
    folderId: quarterFolder,
    folderUrl: `https://drive.google.com/drive/folders/${quarterFolder}`,
    path: [...trail.map((t) => t.name), cfg.driveFolder, result.quarter].join(' / '),
  };
}
