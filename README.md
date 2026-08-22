# Statement Filer

Drop your credit card statements into a web page. It reads them, categorizes every
transaction against your own rules, builds the quarterly Excel workbook, highlights each
statement line in its category, and files the whole lot into the right folder in your
Google Drive.

Everything happens **inside your browser tab**. There is no server, no database, and no
account to sign up for. Your statements are never uploaded anywhere except your own Drive.

---

## What it produces

For a quarter's statements from one card:

```
Annual Audit 2026 Sep
└── Credit Card Statements
    └── Amex
        └── Q3 2026
            ├── Q3 2026 Amex Transactions.xlsx     ← Summary / Data / Review sheets
            ├── Abridged Statements/
            │   ├── Amex Jul 17 2026 Statement.pdf
            │   └── Amex Aug 17 2026 Statement.pdf
            └── Expenses Summary/
                ├── Amex Q3 2026 Business Travel.pdf   ← every Business Travel line
                ├── Amex Q3 2026 Meals.pdf             ←   highlighted in yellow
                ├── Amex Q3 2026 Professional Fees.pdf
                └── … one per category
```

The workbook's Summary sheet carries a **reconciliation block** that compares the total of
everything parsed against the statements' own control totals. If those don't match to the
penny, the app tells you so in red and refuses to file to Drive until it's sorted out.

---

## Setup — three things, once

### 1. Put this in a GitHub repo and turn on Pages

1. On GitHub, create a new repository called `statement-filer`. **Public** — that's what
   makes Pages free.
2. Upload every file and folder from this bundle (or `git push` it).

   Two things are deliberately kept out, and `.gitignore` enforces it: your statement PDFs
   (`samples/`) and your merchant map (`my-rules.json`). Both carry personal data. If you
   ever see either appear in a commit, stop and remove it.
3. Go to **Settings → Pages**, set **Source** to *Deploy from a branch*, branch `main`,
   folder `/ (root)`, and Save.
4. Wait a minute. Your app is live at
   `https://<your-github-username>.github.io/statement-filer/`

Bookmark that. It works from any computer or phone.

### 2. Make a Google OAuth Client ID

This is what lets the page write into your Drive. It takes about ten minutes and you only
do it once.

1. Go to <https://console.cloud.google.com/> and create a project (call it anything).
2. **APIs & Services → Library** → search for **Google Drive API** → **Enable**.
3. **APIs & Services → OAuth consent screen** → choose **External** → fill in an app name
   and your email → Save. Under **Audience → Test users**, add your own Gmail address.
   (While the app is in "Testing", only listed accounts can sign in. That's fine — it's
   just you.)
4. **APIs & Services → Credentials → Create credentials → OAuth client ID**
   - Application type: **Web application**
   - Under **Authorized JavaScript origins**, click *Add URI* and enter
     `https://<your-github-username>.github.io`
     (just the domain — no `/statement-filer` on the end)
   - Create, then copy the **Client ID**. It looks like
     `123456789-abcdef.apps.googleusercontent.com`.

### 3. Paste the Client ID into the app

Open your app, click the **⚙** button top right, paste it in, Save. It's stored in your
browser only. A Client ID is not a password — it's safe in a public repo — but the app
doesn't put it there for you, so each browser you use needs it pasted once.

---

## Using it

### Filing a new quarter

1. Open the app, click **Connect Google Drive** (first time each session).
2. Drag your statement PDFs onto the drop zone. **One card at a time** — the app files by
   card, so mixing Amex and CIBC in one run is refused.
3. Click **Analyze**. You'll see the category summary and the reconciliation result.
4. If it reconciles, click **File everything to Google Drive**.

### Reviewing what it wasn't sure about

Anything the rules can't settle is coded `Review` on the Data sheet and listed on the
**Review** sheet, with a suggested category and a note explaining why it's there. Those
rows are kept out of every other category, so your totals are never quietly wrong.

1. Open the workbook in Drive (or Excel).
2. On the **Review** sheet, put the category you want in the **COMMENTS** column (column F).
   Type the category name exactly, e.g. `Professional Fees`.
3. Save the file.
4. Back in the app, go to **2 · Apply my review**, and either load the file from your
   computer or click **Load the last one from Drive**.
5. It recodes those rows, rebuilds the summary, redraws every highlighted PDF, and — if the
   "remember these decisions" box is ticked — writes them into your rules so the same
   merchants are categorized automatically next quarter.

There's a second block on the Review sheet, *"Categorized, but flagged for a second look"*.
Those already have a category; put something in COMMENTS only if you want to change it.

> The statement PDFs need to be loaded on tab 1 when you apply a review, because the
> highlighted files are redrawn from the originals. If you're coming back a week later,
> just drop the same PDFs in and run Analyze first.

---

## Your rules

### Your merchant list stays off GitHub

This repo is public — that's what makes GitHub Pages free — so the `rules.json` in it ships
with an **empty merchant map**. Merchant names are personal data; a list of where you shop
does not belong in a public repo.

The real map lives in **`my-rules.json`**, which you keep on your own machine. Open the app,
go to the **Rules** tab, click **Import a rules file…**, and pick it. It's stored in that
browser from then on and is never uploaded anywhere. Do it once per browser you use.

If you skip the import, nothing breaks — the app just puts more rows on the Review sheet
at first and learns your answers as you go.

### What's in the rules

`rules.json` is the brain. The **Rules** tab shows it and lets you edit it. It holds:

| Section | What it does |
|---|---|
| `merchants` | Exact statement description → category, per card. Seeded from every workbook you'd already built, with your most recent decision winning where past quarters disagreed. |
| `alwaysReview` | Merchants to put in front of you. A plain entry only fires if you haven't already ruled on that merchant; add `"force": true` to be asked every quarter. |
| `spendCategoryDefaults` | For merchants seen for the first time, falls back to the bank's own spend category (CIBC prints one on every line). |
| `gtaRule` | Meals outside the GTA become Business Travel. Two lists of place names decide which is which — add a city to whichever list it belongs in. |
| `largeAmountReview` | A charge over this amount from a merchant never seen before goes to Review even if a fallback rule would have caught it. |
| `cards` | Per-card file naming, folder names and sheet names, so each card's output matches what's already in your audit folder. |

Edits made in the Rules tab live in that browser. Click **Download rules.json** and commit
it to your repo to make them permanent and available everywhere.

### Anything it can't place goes to Review

The app never guesses silently. If a merchant isn't in the rules and the bank's spend
category doesn't settle it, the row goes to Review. If a city isn't in either GTA list, the
meal is left where it is and flagged. That's deliberate — a wrong number that looks
confident is worse than a question.

---

## What's supported

| Card | Status |
|---|---|
| Amex (Aeroplan Reserve) | Working — verified against your Q3 2026 statements |
| CIBC (Costco World Mastercard) | Working — verified against your Q3 2026 statements |
| Canadian Tire, CI Financial, Rogers, TD Bank | Not yet — send a sample statement and it's a small addition |

Both working parsers were checked by running the app against your real Jul and Aug 2026
statements and confirming the output matches the workbooks you'd already signed off — every
category, to the cent, on both cards.

---

## Privacy and access

The app asks Google for permission to see and manage your Drive files. It needs that broad
scope for one reason: to **find your existing** `Annual Audit … Sep` folder. Google's
narrower app-only scope can't see folders it didn't create itself, which would mean filing
everything into a fresh folder instead of yours.

It only ever writes inside `Annual Audit … Sep / Credit Card Statements /`. The access token
lives in the page's memory and is gone when you close the tab.

---

## For developers

No build step. It's plain ES modules; the three libraries it uses are vendored in
`vendor/` so the app has no CDN dependency and works offline.

```
index.html          the whole UI
app.css
rules.json          your editable rules
js/
  main.js           UI wiring
  pipeline.js       run() and applyReview() — the orchestration
  rules.js          categorization engine
  workbook.js       ExcelJS build + read-back
  highlight.js      pdf-lib category PDFs
  drive.js          Google Identity Services + Drive REST
  parsers/
    base.js         pdf.js word/row extraction shared by all parsers
    amex.js
    cibc.js
    registry.js     detection + lookup
vendor/             pdf.js, ExcelJS, pdf-lib
test/               Playwright harness (npm install, then node test/run.mjs)
```

### Adding a card

Write `js/parsers/<card>.js` exporting `id`, `label`, `detect(pages)`, `parse(pages)` and
optionally `highlightSpan(...)`; add it to `registry.js`; add a `cards.<id>` block to
`rules.json`. `parse` returns `{card, statementLabel, statementDate, controlTotal,
transactions[], transactionPages[]}` where each transaction carries `{date, desc, amount,
page, y0, y1}` — `page`/`y0`/`y1` are what the highlighter draws from.

### Tests

```bash
npm install
node test/run.mjs         # both cards reconcile and match the expected category totals
node test/roundtrip.mjs   # the review loop, end to end
node test/artifacts.mjs   # writes real output files to test/out/ for eyeballing
```

They drive the real app in headless Chromium against the sample statements in `samples/`.
