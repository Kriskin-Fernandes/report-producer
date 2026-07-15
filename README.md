# Merchant Duplicates Report Producer

A **100% client-side** static website that turns a flagged-duplicates account
export (CSV) into a sectioned review report (CSV). Built to be hosted on
**GitHub Pages**.

> **Privacy:** the uploaded file is read and processed entirely in your
> browser using the `FileReader` API. Nothing is uploaded, stored, logged, or
> transmitted. The page ships a Content-Security-Policy with `connect-src
> 'none'`, so the browser itself blocks any outbound network request. There
> are **no third-party scripts or CDNs** — you can save the page and run it
> offline.

---

## What it does

You upload the export CSV (first row = headings, ~840 records below). The tool
groups records by **stripped domain (column U)**, classifies each group, and
produces a report split into three sections:

| Section | Contents |
| --- | --- |
| **Likely duplicates** | Classified groups that are **Low** complexity with **no** relationships and **no** country conflicts on any record. One row per group (the Primary). |
| **Need attention** | Every other classified group. One row per group, plus a **Notes** column flagging relationships / country conflicts. |
| **Unclassified** | Every unclassified record, listed individually (not collapsed to a Primary), grouped with a blank row between groups. |

Download the result as a single CSV.

## Column mapping (read by position / spreadsheet letter)

Columns are read by **position**, matching the spreadsheet letters in the
spec. After you upload, the site shows a **mapping panel** listing the header
it detected at each position so you can confirm the layout matches.

| Report field | Column | Index (0-based) |
| --- | --- | --- |
| Primary Account Name / Account Name | C | 2 |
| Account ID | D | 3 |
| Type | G | 6 |
| Owner | T | 19 |
| Group key (stripped domain) | U | 20 |
| Valid website (Y/N) | W | 22 |
| Classification (Primary/Duplicate/Unclassified) | X | 23 |
| Reason | Y | 24 |
| Complexity | Z | 25 |
| Relationships | AA | 26 |
| Country conflicts | AB | 27 |
| Dupe probability | AC | 28 |

## Report structure (exact output)

```
Likely duplicates
Low complexity, no billing country conflicts, no parent/children
Primary Account Name,Account ID,Type,Owner,Valid website,Reason,Complexity,Dupe probability,Duplicate Account IDs
...one row per qualifying group...
(blank row)
Need attention
Flagged, requires approval
Primary Account Name,Account ID,Type,Owner,Valid website,Reason,Complexity,Dupe probability,Duplicate Account IDs,Notes
...one row per remaining classified group...
(blank row)
Unclassified
Account Name,Account ID,Type,Owner,Valid website,Reason,Complexity,Dupe probability
...every unclassified record, one per row, blank row between groups...
```

### Rules

- **Valid website** (sections 1 & 2): if the group's Primary is `Y` → `Yes`;
  if the Primary is `N` but any Duplicate is `Y` → `Yes-from dupe`; if all are
  `N` → `No`. In the **Unclassified** section each record simply shows its own
  value (`Yes`/`No`).
- **Duplicate Account IDs**: all Account IDs (column D) of the group's
  Duplicate records, in one cell, separated by `; `.
- **Notes** (section 2): `has relationship`, `has country conflict`, or
  `has relationship & country conflict` — based on whether **any** record in
  the group has a non-empty relationships (AA) / conflicts (AB) value.
- A classified group has exactly one **Primary** (used for the displayed row);
  an unclassified group has every record marked **Unclassified**.

## Deploy on GitHub Pages

### Option A — Deploy from this branch (recommended, zero config)

1. Push this repo (already done if you're reading this on GitHub).
2. **Settings → Pages**.
3. **Source:** *Deploy from a branch*.
4. **Branch:** `claude/merchant-duplicates-report-hz4vyo`, folder **`/ (root)`**.
5. Save. Your site publishes at `https://<user>.github.io/<repo>/`.

The files live at the repo root, so no build step is needed. A `.nojekyll`
file is included so Pages serves the files as-is.

### Option B — GitHub Actions

An optional workflow lives at `.github/workflows/deploy.yml`. To use it, set
**Settings → Pages → Source** to *GitHub Actions*. Note: the `github-pages`
environment may restrict deployments to specific branches — if the run fails
with an environment-protection error, either allow this branch in the
environment settings or just use Option A.

## Run the tests

Pure-logic unit tests (Node, no dependencies):

```bash
node test/run.js
```

The logic lives in `report.js` (framework-free, works in the browser and in
Node); `app.js` is the browser-only UI glue.

## Files

```
index.html   # page structure + Content-Security-Policy
styles.css   # styling
report.js    # CSV parse/build/serialize (the logic; testable in Node)
app.js       # browser UI glue (FileReader, preview, download)
test/        # Node test suite + a sample export CSV
```
