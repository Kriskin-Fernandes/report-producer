/*
 * app.js — DOM glue for the report producer.
 *
 * Reads the uploaded CSV locally with FileReader, calls the pure logic in
 * report.js, renders a preview + summary, and offers the result as a download.
 * No network requests are ever made.
 */
(function () {
  'use strict';

  var RP = window.ReportProducer;

  var fileInput = document.getElementById('fileInput');
  var dropzone = document.getElementById('dropzone');
  var fileNameEl = document.getElementById('fileName');
  var errorBox = document.getElementById('errorBox');
  var resultsEl = document.getElementById('results');
  var statsEl = document.getElementById('stats');
  var warningsEl = document.getElementById('warnings');
  var warningList = document.getElementById('warningList');
  var mappingBody = document.querySelector('#mappingTable tbody');
  var previewBody = document.querySelector('#previewTable tbody');
  var previewMore = document.getElementById('previewMore');
  var downloadBtn = document.getElementById('downloadBtn');

  var PREVIEW_ROWS = 60;
  var lastCSV = null;
  var lastDownloadName = 'duplicates-report.csv';

  // ---- File selection wiring ----------------------------------------------
  dropzone.addEventListener('click', function () { fileInput.click(); });
  dropzone.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
  });
  fileInput.addEventListener('change', function () {
    if (fileInput.files && fileInput.files[0]) handleFile(fileInput.files[0]);
  });

  ['dragenter', 'dragover'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) {
      e.preventDefault(); e.stopPropagation();
      dropzone.classList.add('drag');
    });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) {
      e.preventDefault(); e.stopPropagation();
      dropzone.classList.remove('drag');
    });
  });
  dropzone.addEventListener('drop', function (e) {
    var dt = e.dataTransfer;
    if (dt && dt.files && dt.files[0]) handleFile(dt.files[0]);
  });

  downloadBtn.addEventListener('click', function () {
    if (lastCSV == null) return;
    var blob = new Blob([lastCSV], { type: 'text/csv;charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = lastDownloadName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    // Revoke on the next tick so the download has started.
    setTimeout(function () { URL.revokeObjectURL(url); }, 0);
  });

  // ---- Handling a file -----------------------------------------------------
  function handleFile(file) {
    hideError();
    fileNameEl.hidden = false;
    fileNameEl.innerHTML = 'Selected: <strong>' + escapeHtml(file.name) +
      '</strong> (' + formatBytes(file.size) + ')';

    var reader = new FileReader();
    reader.onerror = function () { showError('Could not read the file. Please try again.'); };
    reader.onload = function () {
      try {
        process(String(reader.result), file.name);
      } catch (err) {
        showError((err && err.message) ? err.message : 'Something went wrong while processing the file.');
        resultsEl.hidden = true;
      }
    };
    reader.readAsText(file);
  }

  function process(text, name) {
    var rows = RP.parseCSV(text);
    var result = RP.buildReport(rows);

    lastCSV = RP.toCSV(result.rows);
    lastDownloadName = deriveName(name);

    renderStats(result.stats);
    renderWarnings(result.warnings);
    renderMapping(result.mapping);
    renderPreview(result.rows);

    resultsEl.hidden = false;
    resultsEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // ---- Rendering -----------------------------------------------------------
  function renderStats(s) {
    var items = [
      { cls: 's1', num: s.section1Count, lbl: 'Likely duplicates' },
      { cls: 's2', num: s.section2Count, lbl: 'Need attention' },
      { cls: 's3', num: s.unclassifiedRecordCount, lbl: 'Unclassified records' },
      { cls: '', num: s.totalGroups, lbl: 'Groups total' },
      { cls: '', num: s.totalDataRows, lbl: 'Records read' }
    ];
    statsEl.innerHTML = items.map(function (it) {
      return '<div class="stat ' + it.cls + '">' +
        '<div class="num">' + it.num + '</div>' +
        '<div class="lbl">' + escapeHtml(it.lbl) + '</div></div>';
    }).join('');
  }

  function renderWarnings(warnings) {
    if (!warnings || warnings.length === 0) { warningsEl.hidden = true; return; }
    warningsEl.hidden = false;
    warningList.innerHTML = warnings.map(function (w) {
      return '<li>' + escapeHtml(w) + '</li>';
    }).join('');
  }

  function renderMapping(mapping) {
    mappingBody.innerHTML = mapping.map(function (m) {
      var header = m.header === '' ? '<em class="muted">(empty)</em>' : escapeHtml(m.header);
      return '<tr><td>' + escapeHtml(labelFor(m.field)) + '</td>' +
        '<td>' + escapeHtml(m.letter) + '</td>' +
        '<td>' + m.index + '</td>' +
        '<td>' + header + '</td></tr>';
    }).join('');
  }

  function labelFor(field) {
    var map = {
      primaryName: 'Primary Account Name / Account Name',
      accountId: 'Account ID',
      type: 'Type',
      owner: 'Owner',
      group: 'Group key (stripped domain)',
      validWebsite: 'Valid website (Y/N)',
      classification: 'Classification',
      reason: 'Reason',
      complexity: 'Complexity',
      relationships: 'Relationships',
      conflicts: 'Country conflicts',
      dupeProbability: 'Dupe probability'
    };
    return map[field] || field;
  }

  function renderPreview(rows) {
    var limit = Math.min(rows.length, PREVIEW_ROWS);
    var maxCols = 0;
    for (var i = 0; i < limit; i++) maxCols = Math.max(maxCols, rows[i].length);
    maxCols = Math.max(maxCols, 1);

    var html = '';
    for (var r = 0; r < limit; r++) {
      var row = rows[r];
      var cls = classifyPreviewRow(row);
      if (row.length === 0) {
        html += '<tr class="blank-row"><td colspan="' + maxCols + '"></td></tr>';
        continue;
      }
      if (cls) {
        // Single-cell section/description rows span the whole table.
        html += '<tr class="' + cls + '"><td colspan="' + maxCols + '">' +
          escapeHtml(row[0]) + '</td></tr>';
        continue;
      }
      var tds = '';
      for (var c = 0; c < maxCols; c++) {
        var v = c < row.length ? row[c] : '';
        tds += '<td>' + escapeHtml(v) + '</td>';
      }
      html += '<tr' + (isHeadingRow(row) ? ' class="heading-row"' : '') + '>' + tds + '</tr>';
    }
    previewBody.innerHTML = html;

    if (rows.length > limit) {
      previewMore.hidden = false;
      previewMore.textContent = 'Showing first ' + limit + ' of ' + rows.length +
        ' rows. Download the CSV for the full report.';
    } else {
      previewMore.hidden = true;
    }
  }

  var SECTION_TITLES = { 'Likely duplicates': 1, 'Need attention': 1, 'Unclassified': 1 };
  var SECTION_DESCS = {
    'Low complexity, no billing country conflicts, no parent/children': 1,
    'Flagged, requires approval': 1
  };
  function classifyPreviewRow(row) {
    if (row.length === 1) {
      if (SECTION_TITLES[row[0]]) return 'section-title';
      if (SECTION_DESCS[row[0]]) return 'section-desc';
    }
    return null;
  }
  function isHeadingRow(row) {
    return row[0] === 'Primary Account Name' || row[0] === 'Account Name';
  }

  // ---- Utilities -----------------------------------------------------------
  function deriveName(name) {
    if (!name) return 'duplicates-report.csv';
    var base = name.replace(/\.[^.]+$/, '');
    return (base || 'duplicates') + '-report.csv';
  }
  function formatBytes(n) {
    if (n < 1024) return n + ' B';
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
    return (n / (1024 * 1024)).toFixed(1) + ' MB';
  }
  function showError(msg) { errorBox.hidden = false; errorBox.textContent = msg; }
  function hideError() { errorBox.hidden = true; errorBox.textContent = ''; }
  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
})();
