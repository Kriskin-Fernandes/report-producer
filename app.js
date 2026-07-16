/*
 * app.js — DOM glue for the report producer.
 *
 * Reads the CSV locally, builds the report model (report.js), renders the
 * sheet views + an edit-actions mode, lets the user customize which fields are
 * shown/exported (drag panel), and exports a formatted .xlsx (xlsx.js). Group
 * Actions / Remarks and the chosen Primary are edited in-place on the model and
 * written into the workbook at download time. No network requests are made.
 */
(function () {
  'use strict';

  var RP = window.ReportProducer;
  var XW = window.XlsxWriter;
  var SF = RP.SALESFORCE_BASE;

  var $ = function (id) { return document.getElementById(id); };
  var ACTION_ORDER = ['Merge', 'Ignore', 'Evaluate', 'None'];

  var els = {
    fileInput: $('fileInput'), dropzone: $('dropzone'), fileName: $('fileName'), errorBox: $('errorBox'),
    lockBtn: $('lockBtn'), privacyPop: $('privacyPop'), privacyClose: $('privacyClose'),
    dashCard: $('dashCard'), stats: $('stats'), warnings: $('warnings'), warningList: $('warningList'),
    downloadBtn: $('downloadBtn'),
    viewer: $('viewer'), sheetTabs: $('sheetTabs'), customizeBtn: $('customizeBtn'), editBtn: $('editBtn'),
    sheetTable: $('sheetTable'),
    editOverlay: $('editOverlay'), editSheet: $('editSheet'), editProgress: $('editProgress'),
    editCustomizeBtn: $('editCustomizeBtn'), editClose: $('editClose'),
    editPrimaryName: $('editPrimaryName'), editTags: $('editTags'), editGroupMeta: $('editGroupMeta'),
    editTable: $('editTable'), remarksInput: $('remarksInput'),
    prevGroup: $('prevGroup'), nextGroup: $('nextGroup'), actionButtons: $('actionButtons'),
    customizeOverlay: $('customizeOverlay'), cfTitle: $('cfTitle'), cfClose: $('cfClose'), cfList: $('cfList')
  };

  var state = {
    result: null,
    activeSheet: 0,
    edit: { active: false, sheet: 0, group: 0 },
    cfSheet: 0
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
    var bytes = XW.buildWorkbook(state.result.sheets); // rebuild so edits + customization are included
    var blob = new Blob([bytes], { type: XLSX_MIME });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = lastName;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 0);
  });

  // ---- Viewer controls -----------------------------------------------------
  els.sheetTabs.addEventListener('click', function (e) {
    var t = e.target.closest('.tab'); if (!t) return;
    selectSheet(parseInt(t.getAttribute('data-index'), 10));
  });
  els.editBtn.addEventListener('click', function () { openEdit(state.activeSheet); });
  els.customizeBtn.addEventListener('click', function () { openCustomize(); });
  els.editCustomizeBtn.addEventListener('click', function () { openCustomize(); });

  // ---- Edit-mode controls --------------------------------------------------
  els.editClose.addEventListener('click', closeEdit);
  els.prevGroup.addEventListener('click', function () { navGroup(-1); });
  els.nextGroup.addEventListener('click', function () { navGroup(1); });
  els.remarksInput.addEventListener('input', function () { currentGroup().remarks = els.remarksInput.value; });
  els.actionButtons.addEventListener('click', function (e) {
    var b = e.target.closest('.action-btn'); if (!b) return;
    setAction(b.getAttribute('data-action'));
  });
  els.editTable.addEventListener('change', function (e) {
    var r = e.target.closest('input[type="radio"]'); if (!r) return;
    setPrimary(parseInt(r.getAttribute('data-row'), 10));
  });
  document.addEventListener('keydown', function (e) {
    if (!state.edit.active || !els.customizeOverlay.hidden) return;
    if (e.key === 'Escape') { closeEdit(); return; }
    if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
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
      { cls: 's2', num: s.section2Groups, lbl: 'Problematic-dup groups' },
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

  // ---- Cell rendering (shared) --------------------------------------------
  function td(field, group, row, ri) {
    var val = RP.fieldValue(field, group, row, ri);
    if (field.link && val) {
      var href = SF + encodeURIComponent(val);
      return '<td class="mono"><a href="' + esc(href) + '" target="_blank" rel="noopener noreferrer">' + esc(val) + '</a></td>';
    }
    var cls = field.mono ? 'mono' : '';
    return '<td' + (cls ? ' class="' + cls + '"' : '') + '>' + esc(val) + '</td>';
  }

  function renderSheet() {
    var sheet = state.result.sheets[state.activeSheet];
    var cols = RP.displayedFields(sheet);
    var n = Math.max(cols.length, 1);
    var html = '';
    html += '<tr class="title"><td colspan="' + n + '">' + esc(sheet.title) + '</td></tr>';
    html += '<tr class="desc"><td colspan="' + n + '">' + esc(sheet.description) + '</td></tr>';
    html += '<tr class="head">' + cols.map(function (c) { return '<td>' + esc(c.label) + '</td>'; }).join('') + '</tr>';
    if (!sheet.groups.length) {
      html += '<tr class="empty"><td colspan="' + n + '">No records in this section.</td></tr>';
    } else {
      sheet.groups.forEach(function (g, gi) {
        if (gi > 0) html += '<tr class="blank"><td colspan="' + n + '"></td></tr>';
        g.rows.forEach(function (row, ri) {
          html += '<tr' + (row.isPrimary ? ' class="primary"' : '') + '>' +
            cols.map(function (c) { return td(c, g, row, ri); }).join('') + '</tr>';
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
    els.editOverlay.hidden = false;
    renderEditGroup();
  }

  function closeEdit() {
    state.edit.active = false;
    els.editOverlay.hidden = true;
    renderSheet();
  }

  function navGroup(delta) {
    var sheet = currentSheet();
    var next = state.edit.group + delta;
    if (next < 0 || next >= sheet.groups.length) return;
    state.edit.group = next;
    renderEditGroup();
  }

  function shouldAutoNext(g, sheet) {
    if (sheet.key === 'unclassified') return g.actionChosen && g.rows.some(function (r) { return r.isPrimary; });
    return g.actionChosen;
  }
  function scheduleNext() { setTimeout(function () { navGroup(1); }, 200); }

  function setAction(a) {
    var g = currentGroup();
    g.action = a;
    g.actionChosen = true;
    updateActionButtons();
    if (shouldAutoNext(g, currentSheet())) scheduleNext();
  }

  function setPrimary(idx) {
    var g = currentGroup();
    g.rows.forEach(function (r, i) {
      r.isPrimary = i === idx;
      r.classification = i === idx ? 'Primary' : 'Duplicate';
    });
    var chosen = g.rows[idx];
    g.rows = [chosen].concat(g.rows.filter(function (r) { return r !== chosen; }));
    renderEditGroup();
    // Unclassified auto-advances only once both an action and a primary are set.
    if (currentSheet().key === 'unclassified' && shouldAutoNext(g, currentSheet())) scheduleNext();
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

    // Large centered primary account name.
    var primaryRow = g.rows.filter(function (r) { return r.isPrimary; })[0];
    els.editPrimaryName.textContent = primaryRow
      ? RP.fieldValue({ src: RP.COL.accountName, kind: 'data' }, g, primaryRow, 1) || '(unnamed account)'
      : '⚠ Select a primary account';
    els.editPrimaryName.classList.toggle('no-primary', !primaryRow);

    // Problematic duplicates: show Notes fragments as big centered tags.
    if (sheet.key === 'attention') {
      var frags = (g.notes || '').split(',').map(function (t) { return t.trim(); }).filter(Boolean);
      if (frags.length) {
        els.editTags.hidden = false;
        els.editTags.innerHTML = frags.map(function (t) { return '<span class="edit-tag">' + esc(t) + '</span>'; }).join('');
      } else { els.editTags.hidden = true; els.editTags.innerHTML = ''; }
    } else {
      els.editTags.hidden = true; els.editTags.innerHTML = '';
    }

    els.editGroupMeta.textContent = 'Stripped domain: ' + (g.key || '(blank)') + ' · ' + g.rows.length + ' account(s)';

    // Records table: a Primary radio (replaces Classification) + displayed data
    // fields (no group-level Action/Notes/Remarks).
    var dataCols = RP.displayedFields(sheet).filter(function (f) { return !f.group && f.kind !== 'classification'; });
    var html = '<tr class="head"><td>Primary</td>' +
      dataCols.map(function (c) { return '<td>' + esc(c.label) + '</td>'; }).join('') + '</tr>';
    g.rows.forEach(function (row, ri) {
      var radio = '<td class="pv-radio"><input type="radio" name="editPrimary" data-row="' + ri + '"' +
        (row.isPrimary ? ' checked' : '') + ' aria-label="Set as primary" /></td>';
      html += '<tr' + (row.isPrimary ? ' class="primary"' : '') + '>' + radio +
        dataCols.map(function (c) { return td(c, g, row, ri); }).join('') + '</tr>';
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

  // ---- Customize data (drag to reorder / show / hide) ----------------------
  var dragEl = null;

  function openCustomize() {
    state.cfSheet = state.edit.active ? state.edit.sheet : state.activeSheet;
    els.cfTitle.textContent = 'Customize: ' + state.result.sheets[state.cfSheet].title;
    renderCfList();
    els.customizeOverlay.hidden = false;
  }
  els.cfClose.addEventListener('click', function () { els.customizeOverlay.hidden = true; });

  function cfDivider() { return '<li class="cf-divider" data-divider="1">— display cutoff (shown above · hidden below) —</li>'; }
  function renderCfList() {
    var sheet = state.result.sheets[state.cfSheet];
    var html = '';
    sheet.fields.forEach(function (f, i) {
      if (i === sheet.displayCount) html += cfDivider();
      html += '<li class="cf-item" draggable="true" data-id="' + esc(f.id) + '"><span class="cf-grip" aria-hidden="true">⠿</span>' + esc(f.label) + '</li>';
    });
    if (sheet.displayCount >= sheet.fields.length) html += cfDivider();
    els.cfList.innerHTML = html;
  }

  els.cfList.addEventListener('dragstart', function (e) {
    var it = e.target.closest('.cf-item'); if (!it) return;
    dragEl = it; it.classList.add('dragging');
    if (e.dataTransfer) { e.dataTransfer.effectAllowed = 'move'; try { e.dataTransfer.setData('text/plain', ''); } catch (x) {} }
  });
  els.cfList.addEventListener('dragend', function () {
    if (dragEl) { dragEl.classList.remove('dragging'); dragEl = null; }
    applyCfOrder();
  });
  els.cfList.addEventListener('dragover', function (e) {
    e.preventDefault();
    if (!dragEl) return;
    var after = cfAfterElement(e.clientY);
    if (after == null) els.cfList.appendChild(dragEl);
    else els.cfList.insertBefore(dragEl, after);
  });

  function cfAfterElement(y) {
    var children = Array.prototype.slice.call(els.cfList.children).filter(function (c) { return c !== dragEl; });
    for (var i = 0; i < children.length; i++) {
      var box = children[i].getBoundingClientRect();
      if (y < box.top + box.height / 2) return children[i];
    }
    return null;
  }

  function applyCfOrder() {
    var sheet = state.result.sheets[state.cfSheet];
    var order = [], cutoff = 0, seenDivider = false;
    Array.prototype.forEach.call(els.cfList.children, function (child) {
      if (child.classList.contains('cf-divider')) { seenDivider = true; return; }
      order.push(child.getAttribute('data-id'));
      if (!seenDivider) cutoff++;
    });
    var byId = {};
    sheet.fields.forEach(function (f) { byId[f.id] = f; });
    var reordered = order.map(function (id) { return byId[id]; }).filter(Boolean);
    if (reordered.length === sheet.fields.length) { sheet.fields = reordered; sheet.displayCount = cutoff; }
    if (state.edit.active) renderEditGroup(); else renderSheet();
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
  function showError(msg) { els.errorBox.hidden = false; els.errorBox.textContent = msg; }
  function hideError() { els.errorBox.hidden = true; els.errorBox.textContent = ''; }
  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
})();
