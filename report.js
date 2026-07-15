/*
 * report.js — CSV parsing + report model building.
 *
 * No DOM / browser dependency, so it runs in the browser (via <script>) and
 * under Node.js (for tests). All work is in memory; it never touches the
 * network. It produces a structured "workbook model" (three sheets) that
 * xlsx.js turns into a formatted .xlsx file.
 */
(function (global) {
  'use strict';

  // ---- Column mapping (0-based indexes matching spreadsheet letters) --------
  // A=0 B=1 C=2 D=3 ... G=6 ... T=19 U=20 W=22 X=23 Y=24 Z=25 AA=26 AB=27
  var COL = {
    accountName: 2,      // C  - Account Name (primary or otherwise)
    accountId: 3,        // D  - Account ID (unique per record)
    type: 6,             // G  - Type
    owner: 19,           // T  - Owner
    group: 20,           // U  - Stripped domain (grouping key)
    classification: 23,  // X  - Primary / Duplicate / Unclassified
    reason: 24,          // Y  - Reason
    complexity: 25,      // Z  - Complexity (may carry a "- note" suffix)
    relationships: 26,   // AA - Relationships (parent/children)
    conflicts: 27        // AB - Billing country conflicts
  };

  var COL_LETTERS = {
    accountName: 'C', accountId: 'D', type: 'G', owner: 'T', group: 'U',
    classification: 'X', reason: 'Y', complexity: 'Z',
    relationships: 'AA', conflicts: 'AB'
  };

  // Sheet themes. Fills are HSV(hue, S=20%, V=100%) for blue and peach; gray
  // is a neutral light gray (a 20%-saturated "gray" would not read as gray).
  var THEME = {
    blue: { name: 'blue', headerFill: 'CCE6FF', primaryFill: 'CCE6FF' },
    peach: { name: 'peach', headerFill: 'FFE1CC', primaryFill: 'FFE1CC' },
    gray: { name: 'gray', headerFill: 'E6E6E6', primaryFill: null }
  };

  var ACTION_OPTIONS = ['None', 'Merge', 'Evaluate', 'Ignore'];

  // Column layouts per sheet. `k` is the record field key; `id` flags the
  // monospaced Account ID column.
  var COLS_LIKELY = [
    { h: 'Action', k: 'action' },
    { h: 'Classification', k: 'classification' },
    { h: 'Account Name', k: 'name' },
    { h: 'Account ID', k: 'id', id: true },
    { h: 'Type', k: 'type' },
    { h: 'Owner', k: 'owner' },
    { h: 'Reason', k: 'reason' }
  ];
  var COLS_ATTENTION = COLS_LIKELY.concat([{ h: 'Notes', k: 'notes' }]);
  var COLS_UNCLASSIFIED = [
    { h: 'Action', k: 'action' },
    { h: 'Classification', k: 'classification' },
    { h: 'Account Name', k: 'name' },
    { h: 'Account ID', k: 'id', id: true },
    { h: 'Type', k: 'type' },
    { h: 'Owner', k: 'owner' },
    { h: 'Notes', k: 'notes' }
  ];

  // Fields shown in the "verify your column mapping" panel.
  var USED_FIELDS = ['accountName', 'accountId', 'type', 'owner', 'group',
    'classification', 'reason', 'complexity', 'relationships', 'conflicts'];

  // ---- Small helpers --------------------------------------------------------
  function cell(row, i) { var v = row && row[i]; return v == null ? '' : String(v); }
  function norm(s) { return (s == null ? '' : String(s)).trim(); }
  function isBlank(s) { return norm(s) === ''; }
  function eqi(a, b) { return norm(a).toLowerCase() === String(b).toLowerCase(); }

  // ---- CSV parser (RFC-4180-ish, tolerant) ----------------------------------
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
      if (c === '\r') {
        row.push(field); field = ''; rows.push(row); row = [];
        if (text[i + 1] === '\n') i += 2; else i++;
        continue;
      }
      if (c === '\n') { row.push(field); field = ''; rows.push(row); row = []; i++; continue; }
      field += c; i++;
    }
    row.push(field); rows.push(row);
    if (rows.length) {
      var last = rows[rows.length - 1];
      if (last.length === 1 && last[0] === '') rows.pop();
    }
    return rows;
  }

  function rowIsEmpty(row) {
    for (var i = 0; i < row.length; i++) if (!isBlank(row[i])) return false;
    return true;
  }

  // ---- CSV writer (kept for debugging / tests) ------------------------------
  function toCSV(rows, opts) {
    opts = opts || {};
    function outCell(v) {
      var s = v == null ? '' : String(v);
      return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }
    var body = rows.map(function (r) { return (r || []).map(outCell).join(','); }).join('\r\n');
    var out = body + '\r\n';
    if (opts.bom !== false) out = '﻿' + out;
    return out;
  }

  // ---- Report model builder -------------------------------------------------
  // Returns { sheets, stats, warnings, mapping }.
  function buildReport(allRows) {
    var warnings = [];
    if (!allRows || allRows.length === 0) throw new Error('The file appears to be empty.');

    var header = allRows[0];
    var dataRows = [];
    for (var r = 1; r < allRows.length; r++) {
      if (!rowIsEmpty(allRows[r])) dataRows.push(allRows[r]);
    }
    if (dataRows.length === 0) throw new Error('No data rows found beneath the header row.');

    var mapping = USED_FIELDS.map(function (key) {
      return { field: key, letter: COL_LETTERS[key], index: COL[key], header: norm(cell(header, COL[key])) };
    });

    // Group by stripped domain (column U), preserving first-seen order.
    var groups = [];
    var byKey = Object.create(null);
    var emptyKeyCount = 0;
    for (var d = 0; d < dataRows.length; d++) {
      var rowD = dataRows[d];
      var key = norm(cell(rowD, COL.group));
      if (key === '') emptyKeyCount++;
      var mapKey = ' ' + key;
      var g = byKey[mapKey];
      if (!g) { g = { key: key, records: [] }; byKey[mapKey] = g; groups.push(g); }
      g.records.push(rowD);
    }
    if (emptyKeyCount > 0) {
      warnings.push(emptyKeyCount + ' record(s) have an empty stripped domain (column U); ' +
        'they were grouped together under a blank key.');
    }

    // Classify groups; compute relationship/conflict flags for every group.
    var classifiedGroups = [], unclassifiedGroups = [];
    groups.forEach(function (g) {
      g.hasRelationship = g.records.some(function (rec) { return !isBlank(cell(rec, COL.relationships)); });
      g.hasConflict = g.records.some(function (rec) { return !isBlank(cell(rec, COL.conflicts)); });

      var primaries = g.records.filter(function (rec) { return eqi(cell(rec, COL.classification), 'primary'); });
      var allUnclassified = g.records.every(function (rec) { return eqi(cell(rec, COL.classification), 'unclassified'); });

      if (primaries.length >= 1) {
        g.kind = 'classified';
        g.primary = primaries[0];
        if (primaries.length > 1) {
          warnings.push('Group "' + (g.key || '(blank)') + '" has ' + primaries.length +
            ' records marked "Primary"; used the first one.');
        }
        var comp = norm(cell(g.primary, COL.complexity));
        var dash = comp.indexOf('-');
        g.complexityBase = dash >= 0 ? comp.slice(0, dash).trim() : comp;
        g.complexityExtra = dash >= 0 ? comp.slice(dash + 1).trim() : '';
        classifiedGroups.push(g);
      } else if (allUnclassified) {
        g.kind = 'unclassified';
        unclassifiedGroups.push(g);
      } else {
        g.kind = 'unclassified';
        warnings.push('Group "' + (g.key || '(blank)') + '" has no "Primary" record and is ' +
          'not fully "Unclassified"; placed it in the Unclassified section.');
        unclassifiedGroups.push(g);
      }
    });

    // Split classified groups: Section 1 vs Section 2.
    var section1 = [], section2 = [];
    classifiedGroups.forEach(function (g) {
      if (g.complexityBase.toLowerCase() === 'low' && !g.hasRelationship && !g.hasConflict) {
        section1.push(g);
      } else {
        section2.push(g);
      }
    });

    var sheets = [
      makeSheet('likely', 'Likely duplicates',
        'Low complexity, no billing country conflicts, no parent/children',
        THEME.blue, COLS_LIKELY, section1, 'none'),
      makeSheet('attention', 'Needs Attention', 'Flagged, requires approval',
        THEME.peach, COLS_ATTENTION, section2, 'attention'),
      makeSheet('unclassified', 'Unclassified', 'Records that could not be classified',
        THEME.gray, COLS_UNCLASSIFIED, unclassifiedGroups, 'unclassified')
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

    return { sheets: sheets, stats: stats, warnings: warnings, mapping: mapping };
  }

  function countRecords(groups) {
    return groups.reduce(function (n, g) { return n + g.records.length; }, 0);
  }

  // notesMode: 'none' (no notes), 'attention' (rel/conflict + complexity extra),
  // 'unclassified' (rel/conflict only). Notes sit on the representative row
  // (the Primary, or the first row when there is no primary).
  function makeSheet(key, title, description, theme, columns, groups, notesMode) {
    var groupModels = groups.map(function (g) {
      var noteText = '';
      if (notesMode === 'attention') noteText = buildNotes(g, true);
      else if (notesMode === 'unclassified') noteText = buildNotes(g, false);

      // Primary is always the top row; the rest keep their original order.
      var ordered = g.primary
        ? [g.primary].concat(g.records.filter(function (r) { return r !== g.primary; }))
        : g.records;

      // The group-level note sits on the representative (top) row.
      var rows = ordered.map(function (rec, idx) {
        return {
          action: 'None',
          classification: norm(cell(rec, COL.classification)),
          name: cell(rec, COL.accountName),
          id: cell(rec, COL.accountId),
          type: cell(rec, COL.type),
          owner: cell(rec, COL.owner),
          reason: cell(rec, COL.reason),
          notes: idx === 0 ? noteText : '',
          isPrimary: eqi(cell(rec, COL.classification), 'primary')
        };
      });
      return { rows: rows };
    });

    return {
      key: key, title: title, description: description, theme: theme,
      columns: columns, actionOptions: ACTION_OPTIONS, groups: groupModels
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
    COL: COL, COL_LETTERS: COL_LETTERS, THEME: THEME, ACTION_OPTIONS: ACTION_OPTIONS,
    parseCSV: parseCSV, toCSV: toCSV, buildReport: buildReport
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  global.ReportProducer = api;
})(typeof self !== 'undefined' ? self : (typeof globalThis !== 'undefined' ? globalThis : this));
