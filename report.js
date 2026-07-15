/*
 * report.js — CSV parsing + report model building.
 *
 * No DOM / browser dependency, so it runs in the browser (via <script>) and
 * under Node.js (for tests). All work is in memory; it never touches the
 * network. It produces a structured "workbook model" (three sheets) that
 * xlsx.js turns into a formatted .xlsx file, and that app.js renders and edits.
 */
(function (global) {
  'use strict';

  // ---- Column mapping (0-based indexes matching spreadsheet letters) --------
  // A=0 B=1 C=2 D=3 ... G=6 ... T=19 U=20 V=21 W=22 X=23 Y=24 Z=25 AA=26 AB=27
  var COL = {
    accountName: 2,      // C
    accountId: 3,        // D
    type: 6,             // G
    owner: 19,           // T
    group: 20,           // U  - stripped domain (grouping key; never displayed)
    classification: 23,  // X
    reason: 24,          // Y
    complexity: 25,      // Z  - complexity ("- note" suffix); never displayed
    relationships: 26,   // AA
    conflicts: 27        // AB
  };

  // Columns never shown anywhere (grouping key U, V, complexity Z).
  var NEVER_SHOW = { 20: 1, 21: 1, 25: 1 };

  var THEME = {
    blue: { name: 'blue', headerFill: 'CCE6FF' },
    peach: { name: 'peach', headerFill: 'FFE1CC' },
    gray: { name: 'gray', headerFill: 'E6E6E6' }
  };

  var ACTION_OPTIONS = ['None', 'Merge', 'Evaluate', 'Ignore'];

  // Column defs. `src` = source input-column index (used to know what is
  // "already displayed" for detailed view). Synthetic columns have no src.
  function col(h, k, src, isId) { return { h: h, k: k, src: src, id: !!isId }; }
  var C_ACTION = col('Action', 'action');
  var C_CLASS = col('Classification', 'classification', COL.classification);
  var C_NAME = col('Account Name', 'name', COL.accountName);
  var C_ID = col('Account ID', 'id', COL.accountId, true);
  var C_TYPE = col('Type', 'type', COL.type);
  var C_OWNER = col('Owner', 'owner', COL.owner);
  var C_REASON = col('Reason', 'reason', COL.reason);
  var C_NOTES = col('Notes', 'notes');
  var C_REMARKS = col('Remarks', 'remarks');

  var COLS_LIKELY = [C_ACTION, C_CLASS, C_NAME, C_ID, C_TYPE, C_OWNER, C_REASON, C_REMARKS];
  var COLS_ATTENTION = [C_ACTION, C_CLASS, C_NAME, C_ID, C_TYPE, C_OWNER, C_REASON, C_NOTES, C_REMARKS];
  var COLS_UNCLASSIFIED = [C_ACTION, C_CLASS, C_NAME, C_ID, C_TYPE, C_OWNER, C_NOTES];

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

    // Detail columns for a given sheet's column set: every input column that is
    // not never-shown, not already displayed, and not entirely empty.
    function detailFor(columns) {
      var excluded = Object.create(NEVER_SHOW);
      columns.forEach(function (c) { if (c.src != null) excluded[c.src] = 1; });
      var out = [];
      for (var idx = 0; idx < maxCols; idx++) {
        if (excluded[idx]) continue;
        var h = norm(cell(header, idx));
        var hasData = h !== '' || dataRows.some(function (rr) { return !isBlank(cell(rr, idx)); });
        if (!hasData) continue;
        out.push({ index: idx, h: h !== '' ? h : letterFor(idx) });
      }
      return out;
    }

    var sheets = [
      makeSheet('likely', 'Likely duplicates', 'Low complexity, no billing country conflicts, no parent/children',
        THEME.blue, COLS_LIKELY, section1, 'none', true, detailFor(COLS_LIKELY)),
      makeSheet('attention', 'Needs Attention', 'Flagged, requires approval',
        THEME.peach, COLS_ATTENTION, section2, 'attention', true, detailFor(COLS_ATTENTION)),
      makeSheet('unclassified', 'Unclassified', 'Records that could not be classified',
        THEME.gray, COLS_UNCLASSIFIED, unclassifiedGroups, 'unclassified', false, detailFor(COLS_UNCLASSIFIED))
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
  function makeSheet(key, title, description, theme, columns, groups, notesMode, editable, detailColumns) {
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
          name: cell(rec, COL.accountName),
          id: cell(rec, COL.accountId),
          type: cell(rec, COL.type),
          owner: cell(rec, COL.owner),
          reason: cell(rec, COL.reason),
          isPrimary: eqi(cell(rec, COL.classification), 'primary'),
          raw: rec
        };
      });

      // Action / Remarks / Notes are per-group; they render on the top row.
      // `notes` matches the column key so the writer/UI can read it uniformly.
      return { key: g.key, action: 'None', remarks: '', notes: note, rows: rows };
    });

    return {
      key: key, title: title, description: description, theme: theme,
      columns: columns, detailColumns: detailColumns, actionOptions: ACTION_OPTIONS,
      editable: !!editable, groups: groupModels
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
    parseCSV: parseCSV, toCSV: toCSV, buildReport: buildReport, letterFor: letterFor
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  global.ReportProducer = api;
})(typeof self !== 'undefined' ? self : (typeof globalThis !== 'undefined' ? globalThis : this));
