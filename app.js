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
    oppsInput: $('oppsInput'), oppsBtn: $('oppsBtn'), oppsName: $('oppsName'), oppsClear: $('oppsClear'),
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
    customizeOverlay: $('customizeOverlay'), cfTitle: $('cfTitle'), cfClose: $('cfClose'),
    cfRevert: $('cfRevert'), cfShown: $('cfShown'), cfHidden: $('cfHidden'),
    oppsOverlay: $('oppsOverlay'), oppsTitle: $('oppsTitle'), oppsClose: $('oppsClose'), oppsTable: $('oppsTable')
  };

  var state = {
    result: null,
    activeSheet: 0,
    edit: { active: false, sheet: 0, group: 0 },
    cfSheet: 0,
    cfEditingId: null,       // field id whose name is being edited inline
    opps: null,              // { byAccount: {id: [opp...]}, count, computedAt }
    oppsName: ''
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

  // ---- Opportunities CSV (optional second upload) --------------------------
  els.oppsBtn.addEventListener('click', function () { els.oppsInput.click(); });
  els.oppsInput.addEventListener('change', function () {
    if (els.oppsInput.files && els.oppsInput.files[0]) handleOppsFile(els.oppsInput.files[0]);
  });
  els.oppsClear.addEventListener('click', clearOpps);

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

  // ---- Opportunities popup -------------------------------------------------
  document.addEventListener('click', function (e) {
    var b = e.target.closest('.opps-bubble');
    if (!b || b.classList.contains('zero')) return;
    openOppsPopup(b.getAttribute('data-acct'));
  });
  els.oppsClose.addEventListener('click', function () { els.oppsOverlay.hidden = true; });
  els.oppsOverlay.addEventListener('click', function (e) { if (e.target === els.oppsOverlay) els.oppsOverlay.hidden = true; });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && !els.oppsOverlay.hidden) els.oppsOverlay.hidden = true;
  });

  function openOppsPopup(acct) {
    var list = (state.opps && state.opps.byAccount[acct]) || [];
    els.oppsTitle.textContent = 'Opportunities for ' + acct + ' (' + list.length + ')';
    var head = '<tr class="head"><td>Opportunity ID</td><td>Opportunity owner</td><td>Last modified date</td><td>Owner role</td></tr>';
    var body = list.map(function (o) {
      var link = o.id
        ? '<a href="' + esc(SF + encodeURIComponent(o.id)) + '" target="_blank" rel="noopener noreferrer">' + esc(o.id) + '</a>'
        : '';
      return '<tr><td class="mono">' + link + '</td><td>' + esc(o.owner) + '</td><td>' +
        esc(o.modifiedRaw) + '</td><td>' + esc(o.role) + '</td></tr>';
    }).join('');
    els.oppsTable.innerHTML = head + (body || '<tr class="empty"><td colspan="4">No opportunities.</td></tr>');
    els.oppsOverlay.hidden = false;
  }

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
    if (!state.edit.active || !els.customizeOverlay.hidden || !els.oppsOverlay.hidden) return;
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
    if (state.opps) applyOpps(); // re-attach opportunities to the fresh model
    renderStats(state.result.stats);
    renderWarnings(state.result.warnings);
    buildTabs();
    selectSheet(0);
    els.dashCard.hidden = false;
    els.viewer.hidden = false;
    els.viewer.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  // ---- Opportunities: parse, index, attach --------------------------------
  var OPP = { id: 0, owner: 2, role: 3, account: 5, modified: 23 }; // cols A,C,D,F,X

  function handleOppsFile(file) {
    var reader = new FileReader();
    reader.onerror = function () { showError('Could not read the Opportunities file.'); };
    reader.onload = function () {
      try {
        var rows = RP.parseCSV(String(reader.result));
        state.opps = indexOpps(rows);
        state.oppsName = file.name;
        els.oppsName.hidden = false;
        els.oppsName.innerHTML = 'Opportunities: <strong>' + esc(file.name) + '</strong> (' + state.opps.count + ' linked to ' + Object.keys(state.opps.byAccount).length + ' account(s))';
        els.oppsClear.hidden = false;
        if (state.result) { applyOpps(); rerenderActive(); }
      } catch (err) {
        showError((err && err.message) || 'Could not process the Opportunities file.');
      }
    };
    reader.readAsText(file);
  }

  function clearOpps() {
    state.opps = null; state.oppsName = '';
    els.oppsInput.value = '';
    els.oppsName.hidden = true; els.oppsName.innerHTML = '';
    els.oppsClear.hidden = true;
    if (state.result) {
      state.result.sheets.forEach(removeOppsField);
      state.result.sheets.forEach(function (s) {
        s.groups.forEach(function (g) { g.rows.forEach(function (r) { r.opps = null; r.oppsSummary = ''; }); });
      });
      rerenderActive();
    }
  }

  // Index opportunity rows by Account ID; sort each account's list by last
  // modified date, most recent first.
  function indexOpps(rows) {
    var byAccount = Object.create(null), count = 0;
    var now = new Date();
    for (var i = 1; i < rows.length; i++) { // skip header
      var r = rows[i];
      if (!r || r.every(function (c) { return String(c == null ? '' : c).trim() === ''; })) continue;
      var acct = String(r[OPP.account] == null ? '' : r[OPP.account]).trim();
      if (acct === '') continue;
      var modifiedRaw = String(r[OPP.modified] == null ? '' : r[OPP.modified]).trim();
      var opp = {
        id: String(r[OPP.id] == null ? '' : r[OPP.id]).trim(),
        owner: String(r[OPP.owner] == null ? '' : r[OPP.owner]).trim(),
        role: String(r[OPP.role] == null ? '' : r[OPP.role]).trim(),
        account: acct,
        modifiedRaw: modifiedRaw,
        modified: parseMDY(modifiedRaw)
      };
      (byAccount[acct] || (byAccount[acct] = [])).push(opp);
      count++;
    }
    Object.keys(byAccount).forEach(function (k) {
      byAccount[k].sort(function (a, b) {
        var ta = a.modified ? a.modified.getTime() : -Infinity;
        var tb = b.modified ? b.modified.getTime() : -Infinity;
        return tb - ta;
      });
    });
    return { byAccount: byAccount, count: count, computedAt: now };
  }

  // Attach an Opportunities column + per-row summaries to every sheet.
  function applyOpps() {
    if (!state.result || !state.opps) return;
    var now = state.opps.computedAt;
    state.result.sheets.forEach(function (sheet) {
      ensureOppsField(sheet);
      sheet.groups.forEach(function (g) {
        g.rows.forEach(function (row) {
          var acct = row.raw ? String(row.raw[RP.COL.accountId] == null ? '' : row.raw[RP.COL.accountId]).trim() : '';
          var list = acct ? state.opps.byAccount[acct] : null;
          row.opps = list || [];
          row.oppsSummary = oppsSummary(list, now);
        });
      });
    });
  }

  function ensureOppsField(sheet) {
    if (sheet.fields.some(function (f) { return f.id === 'opps'; })) return;
    var f = RP.makeField('opps', 'Opportunities', 'opportunities', { align: 'center', width: 22 });
    // Insert just before Remarks if it's displayed, else at the end of the
    // displayed region.
    var remIdx = -1;
    sheet.fields.forEach(function (x, i) { if (x.id === 'remarks' && remIdx < 0) remIdx = i; });
    var at = (remIdx >= 0 && remIdx < sheet.displayCount) ? remIdx : sheet.displayCount;
    sheet.fields.splice(at, 0, f);
    sheet.displayCount += 1;
  }

  function removeOppsField(sheet) {
    var idx = -1;
    sheet.fields.forEach(function (x, i) { if (x.id === 'opps' && idx < 0) idx = i; });
    if (idx < 0) return;
    sheet.fields.splice(idx, 1);
    if (idx < sheet.displayCount) sheet.displayCount -= 1;
  }

  function oppsSummary(list, now) {
    if (!list || !list.length) return '0';
    var latest = null;
    list.forEach(function (o) { if (o.modified && (!latest || o.modified > latest)) latest = o.modified; });
    return  list.length + ' (' + (latest ? timeAgo(latest, now) : 'date unknown') + ')';
  }

  function parseMDY(s) {
    s = String(s == null ? '' : s).trim();
    if (!s) return null;
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m) {
      var yr = +m[3]; if (yr < 100) yr += 2000;
      var d = new Date(yr, (+m[1]) - 1, +m[2]);
      return isNaN(d.getTime()) ? null : d;
    }
    var d2 = new Date(s);
    return isNaN(d2.getTime()) ? null : d2;
  }

  function timeAgo(date, now) {
    var ms = now.getTime() - date.getTime();
    if (ms < 0) ms = 0;
    var day = Math.floor(ms / 86400000);
    var yr = Math.floor(day / 365), mo = Math.floor(day / 30), wk = Math.floor(day / 7);
    var hr = Math.floor(ms / 3600000), min = Math.floor(ms / 60000);
    if (yr >= 1) return yr + (yr > 1 ? ' years' : ' year') + ' ago';
    if (mo >= 1) return mo + (mo > 1 ? ' months' : ' month') + ' ago';
    if (wk >= 1) return wk + (wk > 1 ? ' weeks' : ' week') + ' ago';
    if (day >= 1) return day + (day > 1 ? ' days' : ' day') + ' ago';
    if (hr >= 1) return hr + (hr > 1 ? ' hours' : ' hour') + ' ago';
    if (min >= 1) return min + (min > 1 ? ' minutes' : ' minute') + ' ago';
    return 'just now';
  }

  function rerenderActive() {
    if (state.edit.active) renderEditGroup(); else renderSheet();
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
    if (field.kind === 'opportunities') return oppsCell(row);
    var val = RP.fieldValue(field, group, row, ri);
    if (field.link && val) {
      var href = SF + encodeURIComponent(val);
      return '<td class="mono"><a href="' + esc(href) + '" target="_blank" rel="noopener noreferrer">' + esc(val) + '</a></td>';
    }
    var cls = field.mono ? 'mono' : '';
    return '<td' + (cls ? ' class="' + cls + '"' : '') + '>' + esc(val) + '</td>';
  }

  // Opportunities cell: a bubble with "<count> - <time-ago>". Non-zero bubbles
  // are clickable and open a popup listing the account's opportunities.
  function oppsCell(row) {
    var list = (row && row.opps) || [];
    var summary = (row && row.oppsSummary) || '0';
    if (!list.length) return '<td class="opps-cell"><span class="opps-bubble zero">0</span></td>';
    var acct = row.raw ? String(row.raw[RP.COL.accountId] == null ? '' : row.raw[RP.COL.accountId]).trim() : '';
    return '<td class="opps-cell"><button type="button" class="opps-bubble" data-acct="' + esc(acct) +
      '" title="View opportunities">' + esc(summary) + '</button></td>';
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
    // Leaving the group: apply the primary-to-top reorder offscreen.
    RP.reorderPrimaryTop(currentGroup());
    state.edit.active = false;
    els.editOverlay.hidden = true;
    renderSheet();
  }

  function navGroup(delta) {
    var sheet = currentSheet();
    var next = state.edit.group + delta;
    if (next < 0 || next >= sheet.groups.length) return;
    // Reorder the group we're leaving offscreen, so its rows show the primary
    // on top next time it's viewed — but never while it's on screen.
    RP.reorderPrimaryTop(currentGroup());
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
    // NOTE: rows are intentionally NOT reordered here. Watching rows swap while
    // the group is on screen is jarring; the primary-to-top reorder happens
    // offscreen (navGroup / closeEdit). Export always puts the primary on top.
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

  // ---- Customize data (two lists: displayed / hidden; drag + rename) -------
  var dragEl = null;

  function openCustomize() {
    state.cfSheet = state.edit.active ? state.edit.sheet : state.activeSheet;
    state.cfEditingId = null;
    els.cfTitle.textContent = 'Customize: ' + state.result.sheets[state.cfSheet].title;
    renderCfLists();
    els.customizeOverlay.hidden = false;
  }
  els.cfClose.addEventListener('click', function () {
    if (state.cfEditingId) commitRename();
    els.customizeOverlay.hidden = true;
  });
  els.cfRevert.addEventListener('click', function () {
    var sheet = state.result.sheets[state.cfSheet];
    sheet.fields.forEach(function (f) { f.label = f.defaultLabel; });
    state.cfEditingId = null;
    renderCfLists();
    rerenderActive();
  });

  function cfItemHtml(f) {
    if (state.cfEditingId === f.id) {
      return '<li class="cf-item editing" data-id="' + esc(f.id) + '">' +
        '<span class="cf-grip" aria-hidden="true">⠿</span>' +
        '<input class="cf-rename-input" type="text" value="' + esc(f.label) + '" aria-label="Field name" autocomplete="off" />' +
        '<button type="button" class="cf-rename-ok" aria-label="Save name" title="Save">✓</button>' +
        '</li>';
    }
    var renamed = f.label !== f.defaultLabel;
    return '<li class="cf-item" draggable="true" data-id="' + esc(f.id) + '">' +
      '<span class="cf-grip" aria-hidden="true">⠿</span>' +
      '<span class="cf-label' + (renamed ? ' renamed' : '') + '"' +
      (renamed ? ' title="Renamed from ' + esc(f.defaultLabel) + '"' : '') + '>' + esc(f.label) + '</span>' +
      '<button type="button" class="cf-rename" data-id="' + esc(f.id) + '" aria-label="Rename field" title="Rename">✎</button>' +
      '</li>';
  }
  function cfEmpty(msg) { return '<li class="cf-empty">' + esc(msg) + '</li>'; }

  function renderCfLists() {
    var sheet = state.result.sheets[state.cfSheet];
    var shown = sheet.fields.slice(0, sheet.displayCount);
    var hidden = sheet.fields.slice(sheet.displayCount);
    els.cfShown.innerHTML = shown.map(cfItemHtml).join('') || cfEmpty('Drag fields here to show them');
    els.cfHidden.innerHTML = hidden.map(cfItemHtml).join('') || cfEmpty('Drag fields here to hide them');
    if (state.cfEditingId) {
      var inp = els.customizeOverlay.querySelector('.cf-rename-input');
      if (inp) { inp.focus(); inp.select(); }
    }
  }

  // Rename: pencil -> inline input; ✓/Enter/blur commit; Esc cancels.
  els.customizeOverlay.addEventListener('click', function (e) {
    var rn = e.target.closest('.cf-rename');
    if (rn) { if (state.cfEditingId) commitRename(); state.cfEditingId = rn.getAttribute('data-id'); renderCfLists(); return; }
    if (e.target.closest('.cf-rename-ok')) { commitRename(); }
  });
  els.customizeOverlay.addEventListener('keydown', function (e) {
    if (!state.cfEditingId) return;
    if (e.key === 'Enter') { e.preventDefault(); commitRename(); }
    else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); state.cfEditingId = null; renderCfLists(); }
  });
  els.customizeOverlay.addEventListener('blur', function (e) {
    if (state.cfEditingId && e.target && e.target.classList && e.target.classList.contains('cf-rename-input')) commitRename();
  }, true);

  function commitRename() {
    var id = state.cfEditingId; if (!id) return;
    var inp = els.customizeOverlay.querySelector('.cf-rename-input');
    var sheet = state.result.sheets[state.cfSheet];
    var f = sheet.fields.filter(function (x) { return x.id === id; })[0];
    if (f && inp) { var v = inp.value.trim(); f.label = v || f.defaultLabel; }
    state.cfEditingId = null;
    renderCfLists();
    rerenderActive();
  }

  // Drag between/within the two lists.
  [els.cfShown, els.cfHidden].forEach(function (list) {
    list.addEventListener('dragover', function (e) {
      e.preventDefault();
      if (!dragEl) return;
      var empty = list.querySelector('.cf-empty'); if (empty) empty.remove();
      var after = cfAfterElement(list, e.clientY);
      if (after == null) list.appendChild(dragEl);
      else list.insertBefore(dragEl, after);
    });
  });
  els.customizeOverlay.addEventListener('dragstart', function (e) {
    var it = e.target.closest('.cf-item'); if (!it || it.classList.contains('editing')) return;
    dragEl = it; it.classList.add('dragging');
    if (e.dataTransfer) { e.dataTransfer.effectAllowed = 'move'; try { e.dataTransfer.setData('text/plain', ''); } catch (x) {} }
  });
  els.customizeOverlay.addEventListener('dragend', function () {
    if (dragEl) { dragEl.classList.remove('dragging'); dragEl = null; }
    applyCfOrder();
  });

  function cfAfterElement(list, y) {
    var children = Array.prototype.slice.call(list.children).filter(function (c) {
      return c !== dragEl && c.getAttribute('data-id');
    });
    for (var i = 0; i < children.length; i++) {
      var box = children[i].getBoundingClientRect();
      if (y < box.top + box.height / 2) return children[i];
    }
    return null;
  }

  function idsIn(list) {
    return Array.prototype.slice.call(list.children)
      .map(function (c) { return c.getAttribute('data-id'); })
      .filter(Boolean);
  }

  function applyCfOrder() {
    var sheet = state.result.sheets[state.cfSheet];
    var shownIds = idsIn(els.cfShown), hiddenIds = idsIn(els.cfHidden);
    var byId = {};
    sheet.fields.forEach(function (f) { byId[f.id] = f; });
    var order = shownIds.concat(hiddenIds).map(function (id) { return byId[id]; }).filter(Boolean);
    // Guard: keep every field, and never allow an empty displayed list.
    if (order.length === sheet.fields.length && shownIds.length >= 1) {
      sheet.fields = order;
      sheet.displayCount = shownIds.length;
    }
    renderCfLists();
    rerenderActive();
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
