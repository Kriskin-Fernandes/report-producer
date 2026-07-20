# Merchant Duplicates Report Producer

A **100% client-side** static website that turns a flagged-duplicates account
export (CSV) into a formatted **Excel (`.xlsx`) workbook**. Built to be hosted
on **GitHub Pages**.

> **Privacy:** the uploaded file is read and the workbook is built entirely in
> your browser. Nothing is uploaded, logged, or transmitted. The page ships a
> Content-Security-Policy with `connect-src 'none'`, so the browser itself
> blocks any outbound request. There are **no third-party scripts or CDNs** —
> even the `.xlsx` writer is hand-rolled — so you can save the page and run it
> offline. Your in-app edits are saved only in this browser's **local storage**
> (never sent anywhere — see *Saved edits* below).

---

## What it does

You upload the export CSV (first row = headings, records below). The tool
groups records by **stripped domain (column U)**, classifies each group using
the **Classification (column X)** value carried over from the input, and
produces a workbook with **three sheets**, one per section. Every record keeps
its own row (like the input); the **Primary** record in each group is
identified by its Classification value, reordered to the **top** of its group,
and emphasized with bold text.

| Sheet | Theme | Membership | Default columns |
| --- | --- | --- | --- |
| **Likely duplicates** | light blue | Classified groups that are **Low** complexity with **no** relationships and **no** country conflicts (on any record). | Action, Classification, Account Name, Alias account, Account ID, Type, Owner, Account Manager, Total Captured Amount ($USD), Reason, Remarks |
| **Problematic duplicates** | light peach | Every other classified group. | …+ **Notes** (before Remarks) |
| **Unclassified** | light gray | Groups whose records are all Unclassified (plus any group without a clean Primary). | Action, Classification, Account Name, Alias account, Account ID, Type, Owner, Account Manager, Notes, Remarks |

Columns are the *default* — each sheet's shown/exported columns are fully
customizable (see below).

## Using the app

- Four **top-right buttons**: **🌙 / ☀** toggles **dark mode** (remembered on
  this device), and **?** (how-to tutorial), **🍪** (storage & cookie policy)
  and **🔒** (privacy notice) each open a popup.
- The top row holds three cards: **Duplicate accounts CSV** (left 25%), the
  optional **Opportunities CSV** (next 25%), and the **Report** dashboard
  (right 50%). The dashboard shows group counts, records read, and — once an
  Opportunities file is loaded — how many opportunities linked to how many
  accounts, plus how many **milliseconds** the processing took.
- The report table uses the **full screen width**; while you scroll it the
  tab/Customize/Edit buttons stay put — only the rows scroll.
- A **tab selector** switches between full views of each sheet.
- **Account ID** is a clickable link to the account in Salesforce
  (`https://checkout.my.salesforce.com/<id>`), on screen and in the `.xlsx`.
- **Add Opportunities CSV** (optional) loads a second export of opportunities
  and links them to accounts by Account ID (see *Opportunities* below).
- **Customize data** opens a per-sheet panel with two lists — **Displayed** and
  **Hidden**. Drag fields between the lists to show/hide them, or reorder within
  a list; the Displayed list (in order) is exactly what appears in the view
  **and** the downloaded report. The **pencil (✎)** next to a field renames it
  (the new name shows in the view and the `.xlsx`); **Revert names** restores
  the original CSV/built-in names. Every input column is available except
  U, V and Z, which are never shown.
- Each group's first row has an **eye (👁) button**; clicking it opens that
  group directly in **Edit actions**.
- **Edit actions** (all sheets) opens a full-screen mode that steps through the
  groups one at a time. The **primary account name** is the large title; the
  group's accounts are listed with a **Primary radio** to change which account
  is primary; four big buttons (Merge / Ignore / Evaluate / None) assign the
  group's action; and a **Remarks** field adds a note. Choosing an action
  auto-advances to the next group (Unclassified only advances once **both** an
  action and a primary are chosen — **except Ignore**, which needs no primary).
  Groups with **3+ records** show a **Remove (✕)** button per record so records
  that don't belong can be dropped from the group (and from the report). On
  Problematic duplicates the flagged reasons appear as large tags. Use
  **← / →** to move between groups, **Esc** to finish. Actions, remarks,
  removals and primary changes are written into the workbook on download.
  - Changing the primary **does not** re-shuffle the rows while you are looking
    at the group (watching rows jump around is jarring). The primary is moved to
    the top **offscreen** — the next time you open that group, and in the
    exported workbook, which always lists the primary first.

## Opportunities (optional)

If you also have an **Opportunities** export, click **Add Opportunities CSV**.
Opportunities are matched to accounts by **Account ID**; an account may have
zero, one, or many, and opportunities that don't match any account in the
report are ignored. Every table gains an **Opportunities** column showing a
bubble with the **count** and how long ago the **most recently modified**
opportunity was (e.g. `3 (2 weeks ago)`). Clicking a non-empty bubble opens a
popup listing that account's opportunities — **Opportunity ID** (linked to
Salesforce, like Account ID), **Opportunity owner**, **Stage**, **Last modified
date** and **Owner role** — most-recently-modified first. In the downloaded
`.xlsx` the column holds the same summary text.

Opportunities columns (read by position):

| Field | Column | Index (0-based) |
| --- | --- | --- |
| Opportunity ID | A | 0 |
| Opportunity owner | C | 2 |
| Owner role | D | 3 |
| Account ID (links to the account) | F | 5 |
| Stage | N | 13 |
| Last modified date (MM/DD/YYYY) | X | 23 |

## Saved edits (local storage, not cookies)

Your work is remembered so an accidental reload doesn't lose it. When you edit
a report — actions, primary choices, removed records, and the Customize-data
layout (field order, shown/hidden, renames) — the state is written to the
browser's **local storage**, keyed to a fingerprint of the uploaded CSV.
Re-upload the **same file** after a reload and your changes are restored.

This uses local storage, **not cookies**: unlike cookies, local storage is
never attached to any request, so nothing is transmitted (the CSP blocks
outbound requests regardless). It stays on the device until cleared — the
**🍪** popup has **Clear saved edits** (per-file edits) and **Clear all stored
data** (edits **and** the dark-mode preference), or clear site data in your
browser.

## Formatting applied

- **Section header** row: merged across the whole table, 20pt **bold**, centered, Calibri.
- **Section description** row: merged, 15pt *italic*, Calibri.
- **Column headers**: bold.
- **Fill** on the header / description / column-heading rows: `HSV(hue, 20% saturation, 100% value)` — hue per sheet (blue / peach; gray uses a neutral light gray).
- **Primary rows**: **always the top row of their group**, shown in **bold** with no background fill.
- **Borders**: full grid on the header block, a box around each group, and vertical rules between every column (blank row between groups).
- **Account ID** cells are a Salesforce `HYPERLINK` (blue, underlined, monospaced).
- **Notes / Remarks** columns never wrap — each is widened to fit its longest value.
- **Action** is one value **per group**, written on the group's top (primary) row with a real Excel dropdown (`None`, `Merge`, `Evaluate`, `Ignore`); duplicate rows are left blank.
- **Remarks** (from edit mode) are written on the group's top row, in the added Remarks column.
- Top three rows are frozen so headers stay visible while scrolling.

### Notes column

- **Problematic duplicates:** `has relationship` (any record has a value in column AA), `has country conflict` (any record has a value in column AB), plus any text **after the `-`** in the group's Complexity (column Z) — e.g. `Medium - VAT number mismatch` contributes `VAT number mismatch`. Joined with `, `.
- **Unclassified:** relationship / country-conflict flags only.
- The note sits on the group's representative row (the Primary, or the first row when there is no Primary).

> The **Complexity** used for the Low/not-Low membership test is the part
> **before** the `-`. A group whose complexity is `Low - <note>` with no
> relationships or conflicts still lands in **Likely duplicates** (which has no
> Notes column, so the note text is not shown there).

## Column mapping (read by position / spreadsheet letter)

Columns are read by **position**, matching the spreadsheet letters below. If a
real export's column order differs from this layout, the values will land in the
wrong report columns — check your export matches.

| Field | Column | Index (0-based) |
| --- | --- | --- |
| Alias account | B | 1 |
| Account Name | C | 2 |
| Account ID | D | 3 |
| Type | G | 6 |
| Owner | T | 19 |
| Group key (stripped domain) | U | 20 |
| Classification (Primary/Duplicate/Unclassified) | X | 23 |
| Reason | Y | 24 |
| Complexity | Z | 25 |
| Relationships | AA | 26 |
| Country conflicts | AB | 27 |
| Account Manager | AD | 29 |
| Total Captured Amount ($USD) | AE | 30 |

## Deploy on GitHub Pages

### Option A — Deploy from this branch (recommended, zero config)

1. **Settings → Pages**.
2. **Source:** *Deploy from a branch*.
3. **Branch:** `claude/merchant-duplicates-report-hz4vyo`, folder **`/ (root)`**.
4. Save. Your site publishes at `https://<user>.github.io/<repo>/`.

The files live at the repo root, so no build step is needed. A `.nojekyll`
file is included so Pages serves the files as-is.

### Option B — GitHub Actions

An optional workflow lives at `.github/workflows/deploy.yml`. To use it, set
**Settings → Pages → Source** to *GitHub Actions*. (The `github-pages`
environment may restrict deployments to specific branches — if a run fails with
an environment-protection error, allow this branch or use Option A.)

## Run the tests

```bash
node test/run.js
```

This regenerates a sample export, checks the sheet model and the workbook
bytes (valid ZIP + End-Of-Central-Directory), and writes `test/sample-report.xlsx`.
The generated `.xlsx` has been verified to read cleanly under both **SheetJS**
and **ExcelJS** (merges, fonts, fills, borders, monospace IDs, data-validation
dropdowns, frozen panes).

## Files

```
index.html   # page structure + Content-Security-Policy
styles.css   # styling (incl. themed sheet previews)
report.js    # CSV parsing + report/sheet model (testable in Node)
xlsx.js      # hand-rolled, dependency-free XLSX (OOXML/ZIP) writer
app.js       # browser UI glue (FileReader upload, preview, local download)
test/        # Node test suite + sample generator
```
