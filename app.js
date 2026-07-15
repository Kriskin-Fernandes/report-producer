/*
 * app.js — DOM glue for the report producer.
 *
 * Reads the CSV locally, builds the report model (report.js), renders the
 * sheet views + an edit-actions mode, and exports a formatted .xlsx (xlsx.js).
 * Group Actions and Remarks are edited in-place on the model and written into
 * the workbook at download time. No network requests are ever made.
 */
(function () {
  'use strict';

  var RP = window.ReportProducer;
  var XW = window.XlsxWriter;

  var $ = function (id) { return document.getElementById(id); };
  var GROUP_FIELDS = { action: 1, notes: 1, remarks: 1 };
  var ACTION_ORDER = ['Merge', 'Ignore', 'Evaluate', 'None'];

  var els = {
    fileInput: $('fileInput'), dropzone: $('dropzone'), fileName: $('fileName'), errorBox: $('errorBox'),
    lockBtn: $('lockBtn'), privacyPop: $('privacyPop'), privacyClose: $('privacyClose'),
    dashCard: $('dashCard'), stats: $('stats'), warnings: $('warnings'), warningList: $('warningList'),
    downloadBtn: $('downloadBtn'),
    viewer: $('viewer'), sheetTabs: $('sheetTabs'), detailToggle: $('detailToggle'), editBtn: $('editBtn'),
    sheetTable: $('sheetTable'),
    editOverlay: $('editOverlay'), editSheet: $('editSheet'), editProgress: $('editProgress'),
    editDetailToggle: $('editDetailToggle'), editClose: $('editClose'), editNote: $('editNote'),
    editGroupMeta: $('editGroupMeta'), editTable: $('editTable'), remarksInput: $('remarksInput'),
    prevGroup: $('prevGroup'), nextGroup: $('nextGroup'), actionButtons: $('actionButtons')
  };

  var state = {
    result: null,
    activeSheet: 0,
    detailed: false,
    editDetailed: false,
    edit: { active: false, sheet: 0, group: 0 }
  };

  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var lastName = 'duplicates-report.xlsx';

  // ---- File input ----------------------------------------------------------
  els.dropzone.addEventListener('click', function () { els.fileInput.click(); });
  els.dropzone.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); els.fileInput.click(); }
  });
  els.fileInput.addEventListener('change', function () {
    if (els.fileInput.files && els.fileInput.files[0]) handleFile(els.fileInput.files[0]);
  });
  ['dragenter', 'dragover'].forEach(function (ev) {
    els.dropzone.addEventListener(ev, function (e) { e.preventDefault(); e.stopPropagation(); els.dropzone.classList.add('drag'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    els.dropzone.addEventListener(ev, function (e) { e.preventDefault(); e.stopPropagation(); els.dropzone.classList.remove('drag'); });
  });
  els.dropzone.addEventListener('drop', function (e) {
    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  });

  // ---- Privacy popup -------------------------------------------------------
  els.lockBtn.addEventListener('click', function (e) {
    e.stopPropagation();
    var show = els.privacyPop.hidden;
    els.privacyPop.hidden = !show;
    els.lockBtn.setAttribute('aria-expanded', String(show));
  });
  els.privacyClose.addEventListener('click', function () {
    els.privacyPop.hidden = true; els.lockBtn.setAttribute('aria-expanded', 'false');
  });
  document.addEventListener('click', function (e) {
    if (els.privacyPop.hidden) return;
    if (!els.privacyPop.contains(e.target) && e.target !== els.lockBtn) {
      els.privacyPop.hidden = true; els.lockBtn.setAttribute('aria-expanded', 'false');
    }
  });

  // ---- Download ------------------------------------------------------------
  els.downloadBtn.addEventListener('click', function () {
    if (!state.result) return;
    var bytes = XW.buildWorkbook(state.result.sheets); // rebuild so edits are included
    var blob = new Blob([bytes], { type: XLSX_MIME });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = lastName;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 0);
  });

  // ---- Viewer controls -----------------------------------------------------
  els.detailToggle.addEventListener('change', function () { state.detailed = els.detailToggle.checked; renderSheet(); });
  els.sheetTabs.addEventListener('click', function (e) {
    var t = e.target.closest('.tab'); if (!t) return;
    selectSheet(parseInt(t.getAttribute('data-index'), 10));
  });
  els.editBtn.addEventListener('click', function () { openEdit(state.activeSheet); });

  // ---- Edit-mode controls --------------------------------------------------
  els.editClose.addEventListener('click', closeEdit);
  els.prevGroup.addEventListener('click', function () { navGroup(-1); });
  els.nextGroup.addEventListener('click', function () { navGroup(1); });
  els.editDetailToggle.addEventListener('change', function () { state.editDetailed = els.editDetailToggle.checked; renderEditGroup(); });
  els.remarksInput.addEventListener('input', function () { currentGroup().remarks = els.remarksInput.value; });
  els.actionButtons.addEventListener('click', function (e) {
    var b = e.target.closest('.action-btn'); if (!b) return;
    setAction(b.getAttribute('data-action'));
  });
  document.addEventListener('keydown', function (e) {
    if (!state.edit.active) return;
    if (e.key === 'Escape') { closeEdit(); return; }
    if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
    // While typing a remark, let arrows move the cursor until it reaches the
    // text boundary; then they navigate between groups.
    if (e.target === els.remarksInput) {
      var el = els.remarksInput;
      var atEnd = el.selectionStart === el.value.length && el.selectionStart === el.selectionEnd;
      var atStart = el.selectionStart === 0 && el.selectionEnd === 0;
      if (e.key === 'ArrowRight' && !atEnd) return;
      if (e.key === 'ArrowLeft' && !atStart) return;
    }
    e.preventDefault();
    navGroup(e.key === 'ArrowRight' ? 1 : -1);
  });

  // ---- Processing ----------------------------------------------------------
  function handleFile(file) {
    hideError();
    els.fileName.hidden = false;
    els.fileName.innerHTML = 'Selected: <strong>' + esc(file.name) + '</strong> (' + formatBytes(file.size) + ')';
    var reader = new FileReader();
    reader.onerror = function () { showError('Could not read the file. Please try again.'); };
    reader.onload = function () {
      try { process(String(reader.result), file.name); }
      catch (err) { showError((err && err.message) || 'Something went wrong while processing the file.'); }
    };
    reader.readAsText(file);
  }

  function process(text, name) {
    state.result = RP.buildReport(RP.parseCSV(text));
    state.activeSheet = 0;
    lastName = deriveName(name);

    renderStats(state.result.stats);
    renderWarnings(state.result.warnings);
    buildTabs();
    selectSheet(0);

    els.dashCard.hidden = false;
    els.viewer.hidden = false;
    els.viewer.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  function renderStats(s) {
    var items = [
      { cls: 's1', num: s.section1Groups, lbl: 'Likely-dup groups' },
      { cls: 's2', num: s.section2Groups, lbl: 'Needs-attention groups' },
      { cls: 's3', num: s.unclassifiedGroups, lbl: 'Unclassified groups' },
      { cls: '', num: s.totalGroups, lbl: 'Groups total' },
      { cls: '', num: s.totalDataRows, lbl: 'Records read' }
    ];
    els.stats.innerHTML = items.map(function (it) {
      return '<div class="stat ' + it.cls + '"><div class="num">' + it.num + '</div><div class="lbl">' + esc(it.lbl) + '</div></div>';
    }).join('');
  }

  function renderWarnings(warnings) {
    if (!warnings || !warnings.length) { els.warnings.hidden = true; return; }
    els.warnings.hidden = false;
    els.warningList.innerHTML = warnings.map(function (w) { return '<li>' + esc(w) + '</li>'; }).join('');
  }

  function buildTabs() {
    els.sheetTabs.innerHTML = state.result.sheets.map(function (s, i) {
      return '<button type="button" role="tab" class="tab" data-index="' + i + '">' +
        esc(s.title) + ' <span class="tab-count">' + s.groups.length + '</span></button>';
    }).join('');
  }

  function selectSheet(i) {
    state.activeSheet = i;
    var sheet = state.result.sheets[i];
    Array.prototype.forEach.call(els.sheetTabs.children, function (tab, idx) {
      var on = idx === i;
      tab.classList.toggle('active', on);
      tab.setAttribute('aria-selected', String(on));
    });
    els.editBtn.hidden = !sheet.editable;
    els.editBtn.disabled = !sheet.editable || sheet.groups.length === 0;
    renderSheet();
  }

  // Build the combined column list for a sheet view.
  function viewColumns(sheet, detailed) {
    var cols = sheet.columns.slice();
    if (detailed) cols = cols.concat(sheet.detailColumns.map(function (d) { return { h: d.h, detailIndex: d.index }; }));
    return cols;
  }

  function cellValue(col, group, rec, ri) {
    if (col.detailIndex != null) return norm(rec.raw[col.detailIndex]);
    if (GROUP_FIELDS[col.k]) return ri === 0 ? (group[col.k] || '') : '';
    return rec[col.k] == null ? '' : String(rec[col.k]);
  }

  function td(col, group, rec, ri) {
    var cls = col.id ? 'mono' : (col.detailIndex != null ? 'detail' : '');
    return '<td' + (cls ? ' class="' + cls + '"' : '') + '>' + esc(cellValue(col, group, rec, ri)) + '</td>';
  }

  function renderSheet() {
    var sheet = state.result.sheets[state.activeSheet];
    var cols = viewColumns(sheet, state.detailed);
    var n = cols.length;
    var html = '';
    html += '<tr class="title"><td colspan="' + n + '">' + esc(sheet.title) + '</td></tr>';
    html += '<tr class="desc"><td colspan="' + n + '">' + esc(sheet.description) + '</td></tr>';
    html += '<tr class="head">' + cols.map(function (c) { return '<td>' + esc(c.h) + '</td>'; }).join('') + '</tr>';
    if (!sheet.groups.length) {
      html += '<tr class="empty"><td colspan="' + n + '">No records in this section.</td></tr>';
    } else {
      sheet.groups.forEach(function (g, gi) {
        if (gi > 0) html += '<tr class="blank"><td colspan="' + n + '"></td></tr>';
        g.rows.forEach(function (rec, ri) {
          html += '<tr' + (rec.isPrimary ? ' class="primary"' : '') + '>' +
            cols.map(function (c) { return td(c, g, rec, ri); }).join('') + '</tr>';
        });
      });
    }
    els.sheetTable.className = 'pv theme-' + sheet.theme.name;
    els.sheetTable.innerHTML = html;
  }

  // ---- Edit-actions mode ---------------------------------------------------
  function currentSheet() { return state.result.sheets[state.edit.sheet]; }
  function currentGroup() { return currentSheet().groups[state.edit.group]; }

  function openEdit(sheetIndex) {
    var sheet = state.result.sheets[sheetIndex];
    if (!sheet.editable || !sheet.groups.length) return;
    state.edit = { active: true, sheet: sheetIndex, group: 0 };
    els.editDetailToggle.checked = state.editDetailed;
    els.editOverlay.hidden = false;
    renderEditGroup();
  }

  function closeEdit() {
    state.edit.active = false;
    els.editOverlay.hidden = true;
    renderSheet(); // reflect updated actions / remarks
  }

  function navGroup(delta) {
    var sheet = currentSheet();
    var next = state.edit.group + delta;
    if (next < 0 || next >= sheet.groups.length) return;
    state.edit.group = next;
    renderEditGroup();
  }

  function setAction(a) {
    currentGroup().action = a;
    updateActionButtons();
  }

  function updateActionButtons() {
    var action = currentGroup().action;
    Array.prototype.forEach.call(els.actionButtons.children, function (b) {
      b.classList.toggle('active', b.getAttribute('data-action') === action);
    });
  }

  function renderEditGroup() {
    var sheet = currentSheet();
    var g = currentGroup();
    var idx = state.edit.group;

    els.editSheet.textContent = 'Editing: ' + sheet.title;
    els.editProgress.textContent = 'Group ' + (idx + 1) + ' of ' + sheet.groups.length;

    // Notes context (Needs Attention only).
    if (sheet.key === 'attention') {
      els.editNote.hidden = false;
      els.editNote.innerHTML = '<strong>Notes:</strong>' + esc(g.notes || '(none)');
    } else {
      els.editNote.hidden = true;
    }

    els.editGroupMeta.textContent = 'Stripped domain: ' + (g.key || '(blank)') + ' · ' + g.rows.length + ' account(s)';

    // Records table: identity + detail columns (no Action/Notes/Remarks).
    var cols = sheet.columns.filter(function (c) { return !GROUP_FIELDS[c.k]; });
    if (state.editDetailed) cols = cols.concat(sheet.detailColumns.map(function (d) { return { h: d.h, detailIndex: d.index }; }));
    var html = '<tr class="head">' + cols.map(function (c) { return '<td>' + esc(c.h) + '</td>'; }).join('') + '</tr>';
    g.rows.forEach(function (rec, ri) {
      html += '<tr' + (rec.isPrimary ? ' class="primary"' : '') + '>' +
        cols.map(function (c) { return td(c, g, rec, ri); }).join('') + '</tr>';
    });
    els.editTable.className = 'pv theme-' + sheet.theme.name;
    els.editTable.innerHTML = html;

    els.remarksInput.value = g.remarks || '';

    els.actionButtons.innerHTML = ACTION_ORDER.map(function (a) {
      return '<button type="button" class="action-btn" data-action="' + a + '">' + esc(a) + '</button>';
    }).join('');
    updateActionButtons();

    els.prevGroup.disabled = idx === 0;
    els.nextGroup.disabled = idx === sheet.groups.length - 1;
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
  function norm(s) { return (s == null ? '' : String(s)).trim(); }
  function showError(msg) { els.errorBox.hidden = false; els.errorBox.textContent = msg; }
  function hideError() { els.errorBox.hidden = true; els.errorBox.textContent = ''; }
  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
})();
