/*
 * report.js — CSV parsing + report model building.
 *
 * No DOM / browser dependency, so it runs in the browser (via <script>) and
 * under Node.js (for tests). All work is in memory; it never touches the
 * network. It produces a structured "workbook model" (three sheets) that
 * xlsx.js turns into a formatted .xlsx file, and that app.js renders and edits.
 *
 * Fields: each sheet owns an ordered `fields` universe (synthetic + every
 * input data column except U/V/Z) and a `displayCount` cutoff. Fields above the
 * cutoff are shown on screen and exported (in order); the rest are hidden. The
 * customize-data panel reorders this list and moves the cutoff.
 */
(function (global) {
  'use strict';

  // ---- Column mapping (0-based indexes matching spreadsheet letters) --------
  // A=0 B=1 C=2 D=3 ... G=6 ... T=19 U=20 V=21 W=22 X=23 Y=24 Z=25 AA=26 AB=27
  var COL = {
    alias: 1,            // B
    accountName: 2,      // C
    accountId: 3,        // D
    type: 6,             // G
    owner: 19,           // T
    group: 20,           // U  - stripped domain (grouping key; never displayed)
    classification: 23,  // X
    reason: 24,          // Y
    complexity: 25,      // Z  - complexity ("- note" suffix); never displayed
    relationships: 26,   // AA
    conflicts: 27,       // AB
    accountManager: 29,  // AD
    capturedAmount: 30   // AE
  };

  // Columns never shown anywhere (grouping key U, V, complexity Z).
  var NEVER_SHOW = { 20: 1, 21: 1, 25: 1 };

  var THEME = {
    blue: { name: 'blue', headerFill: 'CCE6FF' },
    peach: { name: 'peach', headerFill: 'FFE1CC' },
    gray: { name: 'gray', headerFill: 'E6E6E6' }
  };

  var ACTION_OPTIONS = ['None', 'Merge', 'Evaluate', 'Ignore'];
  var SALESFORCE_BASE = 'https://checkout.my.salesforce.com/';

  // ---- Field descriptors ----------------------------------------------------
  // kind: 'action' | 'notes' | 'remarks' (group-level) | 'classification' | 'id'
  //       | 'data'. Style hints (mono/link/wrap/fit/align/width) drive both the
  // on-screen table and the .xlsx.
  function makeField(id, label, kind, o) {
    o = o || {};
    return {
      id: id, label: label,
      // defaultLabel is the original name (from the CSV header or the built-in
      // label); the customize panel can rename `label` and revert to this.
      defaultLabel: o.defaultLabel != null ? o.defaultLabel : label,
      kind: kind,
      src: o.src == null ? null : o.src,
      group: !!o.group, mono: !!o.mono, link: !!o.link,
      wrap: !!o.wrap, fit: !!o.fit,
      align: o.align || 'left', width: o.width || 18
    };
  }
  function cloneField(f) {
    return { id: f.id, label: f.label, defaultLabel: f.defaultLabel, kind: f.kind, src: f.src, group: f.group,
      mono: f.mono, link: f.link, wrap: f.wrap, fit: f.fit, align: f.align, width: f.width };
  }

  function defAction() { return makeField('action', 'Action', 'action', { group: 1, align: 'center', width: 12 }); }
  function defClass() { return makeField('classification', 'Classification', 'classification', { src: COL.classification, align: 'center', width: 14 }); }
  function defName() { return makeField('name', 'Account Name', 'data', { src: COL.accountName, width: 30 }); }
  function defAlias() { return makeField('alias', 'Alias account', 'data', { src: COL.alias, width: 24 }); }
  function defId() { return makeField('id', 'Account ID', 'id', { src: COL.accountId, mono: 1, link: 1, width: 20 }); }
  function defType() { return makeField('type', 'Type', 'data', { src: COL.type, align: 'center', width: 16 }); }
  function defOwner() { return makeField('owner', 'Owner', 'data', { src: COL.owner, width: 20 }); }
  function defManager() { return makeField('manager', 'Account Manager', 'data', { src: COL.accountManager, width: 22 }); }
  function defCaptured() { return makeField('captured', 'Total Captured Amount ($USD)', 'data', { src: COL.capturedAmount, align: 'right', width: 26 }); }
  function defReason() { return makeField('reason', 'Reason', 'data', { src: COL.reason, wrap: 1, width: 42 }); }
  function defNotes() { return makeField('notes', 'Notes', 'notes', { group: 1, fit: 1, width: 42 }); }
  function defRemarks() { return makeField('remarks', 'Remarks', 'remarks', { group: 1, fit: 1, width: 40 }); }

  // Default displayed fields per sheet (Alias sits right of Account Name).
  var DISPLAYED = {
    likely: function () { return [defAction(), defClass(), defName(), defAlias(), defId(), defType(), defOwner(), defManager(), defCaptured(), defReason(), defRemarks()]; },
    attention: function () { return [defAction(), defClass(), defName(), defAlias(), defId(), defType(), defOwner(), defManager(), defCaptured(), defReason(), defNotes(), defRemarks()]; },
    unclassified: function () { return [defAction(), defClass(), defName(), defAlias(), defId(), defType(), defOwner(), defManager(), defCaptured(), defNotes(), defRemarks()]; }
  };

  // ---- Helpers --------------------------------------------------------------
  function cell(row, i) { var v = row && row[i]; return v == null ? '' : String(v); }
  function norm(s) { return (s == null ? '' : String(s)).trim(); }
  function isBlank(s) { return norm(s) === ''; }
  function eqi(a, b) { return norm(a).toLowerCase() === String(b).toLowerCase(); }
  function letterFor(index) {
    var n = index + 1, s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }

  // Resolve a field's display value for a given row (rowIndex within its group).
  function fieldValue(f, group, row, ri) {
    // Opportunities is a per-row summary computed by app.js (from the optional
    // Opportunities CSV) and stashed on the row as `oppsSummary`.
    if (f.kind === 'opportunities') return (row && row.oppsSummary) || '';
    if (f.group) return ri === 0 ? (group[f.id] || '') : '';
    if (f.kind === 'classification') return row.classification || '';
    return norm(row.raw ? row.raw[f.src] : '');
  }

  // Reorder a group's rows so the Primary is the top row (the rest keep their
  // relative order). Used to reorder "offscreen" after the user leaves a group
  // in edit mode, so the on-screen order never jumps while a group is visible.
  function reorderPrimaryTop(group) {
    if (!group || !group.rows || !group.rows.length) return;
    var primary = null;
    for (var i = 0; i < group.rows.length; i++) { if (group.rows[i].isPrimary) { primary = group.rows[i]; break; } }
    if (!primary) return;
    group.rows = [primary].concat(group.rows.filter(function (r) { return r !== primary; }));
  }

  // Fields shown on screen and exported (above the cutoff).
  function displayedFields(sheet) { return sheet.fields.slice(0, sheet.displayCount); }

  // ---- CSV parser -----------------------------------------------------------
  function parseCSV(text) {
    if (text == null) return [];
    text = String(text);
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    var rows = [], row = [], field = '', inQuotes = false, i = 0, n = text.length;
    while (i < n) {
      var c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQuotes = false; i++; continue;
        }
        field += c; i++; continue;
      }
      if (c === '"') { inQuotes = true; i++; continue; }
      if (c === ',') { row.push(field); field = ''; i++; continue; }
      if (c === '\r') { row.push(field); field = ''; rows.push(row); row = []; if (text[i + 1] === '\n') i += 2; else i++; continue; }
      if (c === '\n') { row.push(field); field = ''; rows.push(row); row = []; i++; continue; }
      field += c; i++;
    }
    row.push(field); rows.push(row);
    if (rows.length) { var last = rows[rows.length - 1]; if (last.length === 1 && last[0] === '') rows.pop(); }
    return rows;
  }

  function rowIsEmpty(row) {
    for (var i = 0; i < row.length; i++) if (!isBlank(row[i])) return false;
    return true;
  }

  function toCSV(rows, opts) {
    opts = opts || {};
    function outCell(v) { var s = v == null ? '' : String(v); return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
    var out = rows.map(function (r) { return (r || []).map(outCell).join(','); }).join('\r\n') + '\r\n';
    if (opts.bom !== false) out = '﻿' + out;
    return out;
  }

  // ---- Report model builder -------------------------------------------------
  function buildReport(allRows) {
    var warnings = [];
    if (!allRows || allRows.length === 0) throw new Error('The file appears to be empty.');

    var header = allRows[0];
    var dataRows = [];
    var maxCols = header.length;
    for (var r = 1; r < allRows.length; r++) {
      if (!rowIsEmpty(allRows[r])) { dataRows.push(allRows[r]); if (allRows[r].length > maxCols) maxCols = allRows[r].length; }
    }
    if (dataRows.length === 0) throw new Error('No data rows found beneath the header row.');

    // Group by stripped domain (column U), preserving first-seen order.
    var groups = [], byKey = Object.create(null), emptyKeyCount = 0;
    for (var d = 0; d < dataRows.length; d++) {
      var rowD = dataRows[d];
      var key = norm(cell(rowD, COL.group));
      if (key === '') emptyKeyCount++;
      var mapKey = ' ' + key, g = byKey[mapKey];
      if (!g) { g = { key: key, records: [] }; byKey[mapKey] = g; groups.push(g); }
      g.records.push(rowD);
    }
    if (emptyKeyCount > 0) {
      warnings.push(emptyKeyCount + ' record(s) have an empty stripped domain (column U); grouped under a blank key.');
    }

    // Classify.
    var classifiedGroups = [], unclassifiedGroups = [];
    groups.forEach(function (gr) {
      gr.hasRelationship = gr.records.some(function (rec) { return !isBlank(cell(rec, COL.relationships)); });
      gr.hasConflict = gr.records.some(function (rec) { return !isBlank(cell(rec, COL.conflicts)); });
      var primaries = gr.records.filter(function (rec) { return eqi(cell(rec, COL.classification), 'primary'); });
      var allUnclassified = gr.records.every(function (rec) { return eqi(cell(rec, COL.classification), 'unclassified'); });
      if (primaries.length >= 1) {
        gr.kind = 'classified';
        gr.primary = primaries[0];
        if (primaries.length > 1) warnings.push('Group "' + (gr.key || '(blank)') + '" has ' + primaries.length + ' "Primary" records; used the first.');
        var comp = norm(cell(gr.primary, COL.complexity));
        var dash = comp.indexOf('-');
        gr.complexityBase = dash >= 0 ? comp.slice(0, dash).trim() : comp;
        gr.complexityExtra = dash >= 0 ? comp.slice(dash + 1).trim() : '';
        classifiedGroups.push(gr);
      } else {
        gr.kind = 'unclassified';
        if (!allUnclassified) warnings.push('Group "' + (gr.key || '(blank)') + '" has no "Primary" and is not fully "Unclassified"; placed in Unclassified.');
        unclassifiedGroups.push(gr);
      }
    });

    var section1 = [], section2 = [];
    classifiedGroups.forEach(function (gr) {
      if (gr.complexityBase.toLowerCase() === 'low' && !gr.hasRelationship && !gr.hasConflict) section1.push(gr);
      else section2.push(gr);
    });

    // Hidden data fields = every input column that is not never-shown, not
    // already displayed, and not entirely empty (in first-seen order).
    function hiddenFieldsFor(displayed) {
      var used = Object.create(NEVER_SHOW);
      displayed.forEach(function (f) { if (f.src != null) used[f.src] = 1; });
      var out = [];
      for (var idx = 0; idx < maxCols; idx++) {
        if (used[idx]) continue;
        var h = norm(cell(header, idx));
        var hasData = h !== '' || dataRows.some(function (rr) { return !isBlank(cell(rr, idx)); });
        if (!hasData) continue;
        var label = h !== '' ? h : letterFor(idx);
        out.push(makeField('col' + idx, label, 'data', { src: idx, width: Math.min(40, Math.max(14, label.length + 2)) }));
      }
      return out;
    }

    function fieldsFor(key) {
      var displayed = DISPLAYED[key]();
      var hidden = hiddenFieldsFor(displayed);
      return { fields: displayed.concat(hidden), displayCount: displayed.length };
    }

    var sheets = [
      makeSheet('likely', 'Likely duplicates', 'Possible duplicates with no issues, no billing country conflicts, no parent/children relationships',
        THEME.blue, fieldsFor('likely'), section1, 'none', true),
      makeSheet('attention', 'Problematic duplicates', 'Possible duplicates with problems flagged by the system',
        THEME.peach, fieldsFor('attention'), section2, 'attention', true),
      makeSheet('unclassified', 'Unclassified', 'Records that could not be classified',
        THEME.gray, fieldsFor('unclassified'), unclassifiedGroups, 'unclassified', true)
    ];

    var stats = {
      totalDataRows: dataRows.length,
      totalGroups: groups.length,
      section1Groups: section1.length,
      section2Groups: section2.length,
      unclassifiedGroups: unclassifiedGroups.length,
      section1Records: countRecords(section1),
      section2Records: countRecords(section2),
      unclassifiedRecords: countRecords(unclassifiedGroups)
    };

    return { sheets: sheets, stats: stats, warnings: warnings, header: header };
  }

  function countRecords(groups) { return groups.reduce(function (n, g) { return n + g.records.length; }, 0); }

  // notesMode: 'none' | 'attention' | 'unclassified'.
  function makeSheet(key, title, description, theme, fieldConf, groups, notesMode, editable) {
    var groupModels = groups.map(function (g) {
      var note = '';
      if (notesMode === 'attention') note = buildNotes(g, true);
      else if (notesMode === 'unclassified') note = buildNotes(g, false);

      // Primary is always the top row; the rest keep their original order.
      var ordered = g.primary
        ? [g.primary].concat(g.records.filter(function (r) { return r !== g.primary; }))
        : g.records;

      var rows = ordered.map(function (rec) {
        return {
          classification: norm(cell(rec, COL.classification)),
          isPrimary: eqi(cell(rec, COL.classification), 'primary'),
          raw: rec
        };
      });

      // Action / Remarks / Notes are per-group; they render on the top row.
      // actionChosen tracks whether the user actively picked an action.
      return { key: g.key, action: 'None', actionChosen: false, remarks: '', notes: note, rows: rows };
    });

    return {
      key: key, title: title, description: description, theme: theme,
      fields: fieldConf.fields, displayCount: fieldConf.displayCount,
      actionOptions: ACTION_OPTIONS, editable: !!editable, groups: groupModels
    };
  }

  function buildNotes(g, includeComplexity) {
    var frags = [];
    if (g.hasRelationship) frags.push('has relationship');
    if (g.hasConflict) frags.push('has country conflict');
    if (includeComplexity && g.complexityExtra) frags.push(g.complexityExtra);
    return frags.join(', ');
  }

  var api = {
    COL: COL, THEME: THEME, ACTION_OPTIONS: ACTION_OPTIONS, NEVER_SHOW: NEVER_SHOW,
    SALESFORCE_BASE: SALESFORCE_BASE,
    parseCSV: parseCSV, toCSV: toCSV, buildReport: buildReport, letterFor: letterFor,
    fieldValue: fieldValue, displayedFields: displayedFields, cloneField: cloneField,
    makeField: makeField, reorderPrimaryTop: reorderPrimaryTop
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  global.ReportProducer = api;
})(typeof self !== 'undefined' ? self : (typeof globalThis !== 'undefined' ? globalThis : this));
