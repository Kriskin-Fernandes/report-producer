/*
 * app.js — DOM glue for the report producer.
 *
 * Reads the uploaded CSV locally with FileReader, calls the pure logic in
 * report.js, builds a formatted .xlsx with xlsx.js, renders a preview +
 * summary, and offers the workbook as a local download. No network requests.
 */
(function () {
  'use strict';

  var RP = window.ReportProducer;
  var XW = window.XlsxWriter;

  var fileInput = document.getElementById('fileInput');
  var dropzone = document.getElementById('dropzone');
  var fileNameEl = document.getElementById('fileName');
  var errorBox = document.getElementById('errorBox');
  var resultsEl = document.getElementById('results');
  var statsEl = document.getElementById('stats');
  var warningsEl = document.getElementById('warnings');
  var warningList = document.getElementById('warningList');
  var mappingBody = document.querySelector('#mappingTable tbody');
  var previewsEl = document.getElementById('previews');
  var downloadBtn = document.getElementById('downloadBtn');

  var PREVIEW_RECORD_ROWS = 40; // cap record rows shown per sheet preview
  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  var lastBytes = null;
  var lastDownloadName = 'duplicates-report.xlsx';

  // ---- File selection wiring ----------------------------------------------
  dropzone.addEventListener('click', function () { fileInput.click(); });
  dropzone.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
  });
  fileInput.addEventListener('change', function () {
    if (fileInput.files && fileInput.files[0]) handleFile(fileInput.files[0]);
  });

  ['dragenter', 'dragover'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) { e.preventDefault(); e.stopPropagation(); dropzone.classList.add('drag'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) { e.preventDefault(); e.stopPropagation(); dropzone.classList.remove('drag'); });
  });
  dropzone.addEventListener('drop', function (e) {
    var dt = e.dataTransfer;
    if (dt && dt.files && dt.files[0]) handleFile(dt.files[0]);
  });

  downloadBtn.addEventListener('click', function () {
    if (!lastBytes) return;
    var blob = new Blob([lastBytes], { type: XLSX_MIME });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = lastDownloadName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 0);
  });

  // ---- Handling a file -----------------------------------------------------
  function handleFile(file) {
    hideError();
    fileNameEl.hidden = false;
    fileNameEl.innerHTML = 'Selected: <strong>' + escapeHtml(file.name) + '</strong> (' + formatBytes(file.size) + ')';

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
    var result = RP.buildReport(RP.parseCSV(text));
    lastBytes = XW.buildWorkbook(result.sheets);
    lastDownloadName = deriveName(name);

    renderStats(result.stats);
    renderWarnings(result.warnings);
    renderMapping(result.mapping);
    renderPreviews(result.sheets);

    resultsEl.hidden = false;
    resultsEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // ---- Rendering -----------------------------------------------------------
  function renderStats(s) {
    var items = [
      { cls: 's1', num: s.section1Groups, lbl: 'Likely-dup groups' },
      { cls: 's2', num: s.section2Groups, lbl: 'Needs-attention groups' },
      { cls: 's3', num: s.unclassifiedGroups, lbl: 'Unclassified groups' },
      { cls: '', num: s.totalGroups, lbl: 'Groups total' },
      { cls: '', num: s.totalDataRows, lbl: 'Records read' }
    ];
    statsEl.innerHTML = items.map(function (it) {
      return '<div class="stat ' + it.cls + '"><div class="num">' + it.num +
        '</div><div class="lbl">' + escapeHtml(it.lbl) + '</div></div>';
    }).join('');
  }

  function renderWarnings(warnings) {
    if (!warnings || warnings.length === 0) { warningsEl.hidden = true; return; }
    warningsEl.hidden = false;
    warningList.innerHTML = warnings.map(function (w) { return '<li>' + escapeHtml(w) + '</li>'; }).join('');
  }

  function renderMapping(mapping) {
    mappingBody.innerHTML = mapping.map(function (m) {
      var header = m.header === '' ? '<em class="muted">(empty)</em>' : escapeHtml(m.header);
      return '<tr><td>' + escapeHtml(labelFor(m.field)) + '</td><td>' + escapeHtml(m.letter) +
        '</td><td>' + m.index + '</td><td>' + header + '</td></tr>';
    }).join('');
  }

  function labelFor(field) {
    var map = {
      accountName: 'Account Name', accountId: 'Account ID', type: 'Type', owner: 'Owner',
      group: 'Group key (stripped domain)', classification: 'Classification', reason: 'Reason',
      complexity: 'Complexity', relationships: 'Relationships', conflicts: 'Country conflicts'
    };
    return map[field] || field;
  }

  function renderPreviews(sheets) {
    previewsEl.innerHTML = sheets.map(function (sheet) { return sheetPreview(sheet); }).join('');
  }

  function sheetPreview(sheet) {
    var cols = sheet.columns;
    var nCols = cols.length;
    var totalRecords = sheet.groups.reduce(function (n, g) { return n + g.rows.length; }, 0);

    var rows = '';
    rows += '<tr class="title"><td colspan="' + nCols + '">' + escapeHtml(sheet.title) + '</td></tr>';
    rows += '<tr class="desc"><td colspan="' + nCols + '">' + escapeHtml(sheet.description) + '</td></tr>';
    rows += '<tr class="head">' + cols.map(function (c) { return '<td>' + escapeHtml(c.h) + '</td>'; }).join('') + '</tr>';

    if (totalRecords === 0) {
      rows += '<tr><td colspan="' + nCols + '" class="empty">No records in this section.</td></tr>';
    } else {
      var shown = 0, truncated = false;
      for (var gi = 0; gi < sheet.groups.length && !truncated; gi++) {
        if (gi > 0) rows += '<tr class="blank"><td colspan="' + nCols + '"></td></tr>';
        var grp = sheet.groups[gi];
        for (var ri = 0; ri < grp.rows.length; ri++) {
          if (shown >= PREVIEW_RECORD_ROWS) { truncated = true; break; }
          var rec = grp.rows[ri];
          var tds = cols.map(function (c) {
            var v = rec[c.k];
            v = v == null ? '' : String(v);
            return '<td' + (c.id ? ' class="mono"' : '') + '>' + escapeHtml(v) + '</td>';
          }).join('');
          rows += '<tr' + (rec.isPrimary ? ' class="primary"' : '') + '>' + tds + '</tr>';
          shown++;
        }
      }
      if (truncated) {
        rows += '<tr class="more"><td colspan="' + nCols + '">Showing first ' + PREVIEW_RECORD_ROWS +
          ' records — download the .xlsx for all ' + totalRecords + '.</td></tr>';
      }
    }

    return '<div class="sheet-preview theme-' + sheet.theme.name + '">' +
      '<h4>' + escapeHtml(sheet.title) + ' <span class="muted">· ' + sheet.groups.length +
      ' group(s), ' + totalRecords + ' record(s)</span></h4>' +
      '<div class="table-scroll"><table class="pv">' + rows + '</table></div></div>';
  }

  // ---- Utilities -----------------------------------------------------------
  function deriveName(name) {
    if (!name) return 'duplicates-report.xlsx';
    var base = name.replace(/\.[^.]+$/, '');
    return (base || 'duplicates') + '-report.xlsx';
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
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
})();
