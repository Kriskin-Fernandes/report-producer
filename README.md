# Merchant Duplicates Report Producer

A **100% client-side** static website that turns a flagged-duplicates account
export (CSV) into a formatted **Excel (`.xlsx`) workbook**. Built to be hosted
on **GitHub Pages**.

> **Privacy:** the uploaded file is read and the workbook is built entirely in
> your browser. Nothing is uploaded, stored, logged, or transmitted. The page
> ships a Content-Security-Policy with `connect-src 'none'`, so the browser
> itself blocks any outbound request. There are **no third-party scripts or
> CDNs** — even the `.xlsx` writer is hand-rolled — so you can save the page
> and run it offline.

---

## What it does

You upload the export CSV (first row = headings, records below). The tool
groups records by **stripped domain (column U)**, classifies each group using
the **Classification (column X)** value carried over from the input, and
produces a workbook with **three sheets**, one per section. Every record keeps
its own row (like the input); the **Primary** record in each group is
identified by its Classification value and is emphasized with bold text and a
colored row.

| Sheet | Theme | Membership | Columns |
| --- | --- | --- | --- |
| **Likely duplicates** | light blue | Classified groups that are **Low** complexity with **no** relationships and **no** country conflicts (on any record). | Action, Classification, Account Name, Account ID, Type, Owner, Reason |
| **Needs Attention** | light peach | Every other classified group. | Action, Classification, Account Name, Account ID, Type, Owner, Reason, **Notes** |
| **Unclassified** | light gray | Groups whose records are all Unclassified (plus any group without a clean Primary). | Action, Classification, Account Name, Account ID, Type, Owner, **Notes** |

## Formatting applied

- **Section header** row: merged across the whole table, 20pt **bold**, centered, Calibri.
- **Section description** row: merged, 15pt *italic*, Calibri.
- **Column headers**: bold.
- **Fill** on the header / description / column-heading rows: `HSV(hue, 20% saturation, 100% value)` — hue per sheet (blue / peach; gray uses a neutral light gray).
- **Primary rows**: **always the top row of their group**, shown in **bold** with no background fill.
- **Borders**: full grid on the header block, a box around each group, and vertical rules between every column (blank row between groups).
- **Account ID** cells use a monospaced font (Consolas).
- **Notes** column never wraps — it is widened to fit its longest value so each note stays on a single line.
- **Action** column: a real Excel dropdown (`None`, `Merge`, `Evaluate`, `Ignore`), defaulting to `None`.
- Top three rows are frozen so headers stay visible while scrolling.

### Notes column

- **Needs Attention:** `has relationship` (any record has a value in column AA), `has country conflict` (any record has a value in column AB), plus any text **after the `-`** in the group's Complexity (column Z) — e.g. `Medium - VAT number mismatch` contributes `VAT number mismatch`. Joined with `, `.
- **Unclassified:** relationship / country-conflict flags only.
- The note sits on the group's representative row (the Primary, or the first row when there is no Primary).

> The **Complexity** used for the Low/not-Low membership test is the part
> **before** the `-`. A group whose complexity is `Low - <note>` with no
> relationships or conflicts still lands in **Likely duplicates** (which has no
> Notes column, so the note text is not shown there).

## Column mapping (read by position / spreadsheet letter)

Columns are read by **position**, matching the spreadsheet letters. After you
upload, the site shows a **mapping panel** listing the header it detected at
each position so you can confirm the layout matches.

| Field | Column | Index (0-based) |
| --- | --- | --- |
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
