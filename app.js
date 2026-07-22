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
  var ACTION_ORDER = ['Merge', 'Close', 'Evaluate', 'None'];

  var els = {
    fileInput: $('fileInput'), dropzone: $('dropzone'), fileName: $('fileName'), errorBox: $('errorBox'),
    oppsInput: $('oppsInput'), oppsDropzone: $('oppsDropzone'), oppsName: $('oppsName'), oppsClear: $('oppsClear'),
    lockBtn: $('lockBtn'), privacyPop: $('privacyPop'), themeBtn: $('themeBtn'),
    helpBtn: $('helpBtn'), helpPop: $('helpPop'),
    cookieBtn: $('cookieBtn'), cookiePop: $('cookiePop'), cookieClear: $('cookieClear'), cookieClearAll: $('cookieClearAll'), cookieCleared: $('cookieCleared'),
    dashCard: $('dashCard'), stats: $('stats'), warnings: $('warnings'), warningList: $('warningList'),
    downloadBtn: $('downloadBtn'),
    viewer: $('viewer'), sheetTabs: $('sheetTabs'), customizeBtn: $('customizeBtn'), editBtn: $('editBtn'),
    sheetTable: $('sheetTable'), sheetScroll: $('sheetScroll'),
    reviewView: $('reviewView'), reviewToolbar: $('reviewToolbar'), reviewSearch: $('reviewSearch'), reviewResults: $('reviewResults'),
    editOverlay: $('editOverlay'), editSheet: $('editSheet'), editProgress: $('editProgress'),
    editCustomizeBtn: $('editCustomizeBtn'), editClose: $('editClose'),
    editPrimaryName: $('editPrimaryName'), editTags: $('editTags'), editGroupMeta: $('editGroupMeta'),
    editTable: $('editTable'), editUndo: $('editUndo'), remarksInput: $('remarksInput'),
    prevGroup: $('prevGroup'), nextGroup: $('nextGroup'), actionButtons: $('actionButtons'),
    customizeOverlay: $('customizeOverlay'), cfTitle: $('cfTitle'), cfClose: $('cfClose'),
    cfRevert: $('cfRevert'), cfShown: $('cfShown'), cfHidden: $('cfHidden'), cfPresets: $('cfPresets'),
    tagOverlay: $('tagOverlay'), tagTitle: $('tagTitle'), tagHint: $('tagHint'), tagCurrent: $('tagCurrent'),
    tagInput: $('tagInput'), tagAdd: $('tagAdd'), tagPool: $('tagPool'), tagApply: $('tagApply'), tagClose: $('tagClose'),
    oppsOverlay: $('oppsOverlay'), oppsTitle: $('oppsTitle'), oppsClose: $('oppsClose'), oppsTable: $('oppsTable')
  };

  // ---- Review constants -----------------------------------------------------
  // Five tables, partitioning every group by its assigned action + remarks.
  var REVIEW_TABLES = [
    { title: 'Merge — no remarks' },
    { title: 'Merge — with remarks' },
    { title: 'Evaluate / None with remarks' },
    { title: 'Close' },
    { title: 'None' }
  ];
  var NEXT_STEPS = ['None', 'Contact', 'Follow up', 'Done'];
  var ANY_TAG = 'Any'; // filter-only wildcard (never a real assignable tag)

  var state = {
    result: null,
    activeSheet: 0,          // 0..N-1 = sheets; N = the Review tab
    edit: { active: false, sheet: 0, group: 0 },
    cfTarget: null,          // sheet-like object the customize panel is editing
    cfEditingId: null,       // field id whose name is being edited inline
    opps: null,              // { byAccount: {id: [opp...]}, count, computedAt }
    oppsName: '',
    csvKey: null,            // localStorage key derived from the uploaded CSV
    processMs: null,         // how long parsing + building the report took
    // Review state
    reviewSheet: null,       // the review field universe (result.review)
    reviewInclude: null,     // [bool x5] which partition tables are shown
    reviewSearch: {},        // fieldId -> search text
    tagFilter: {},           // tagName -> true (include; plus ANY_TAG)
    tagExclude: {},          // tagName -> true (exclude)
    tagPool: [],             // global pool of people-tag names
    tagTarget: null,         // group being tag-edited, or '*' for bulk-apply
    tagBulkPick: null        // Set of tags chosen in bulk mode
  };
  var now = (typeof performance !== 'undefined' && performance.now) ? function () { return performance.now(); } : function () { return Date.now(); };

  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var lastName = 'duplicates-report.xlsx';

  // ---- File inputs (drag + click) ------------------------------------------
  function wireDropzone(zone, input, handler) {
    zone.addEventListener('click', function () { input.click(); });
    zone.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
    });
    input.addEventListener('change', function () {
      if (input.files && input.files[0]) handler(input.files[0]);
    });
    ['dragenter', 'dragover'].forEach(function (ev) {
      zone.addEventListener(ev, function (e) { e.preventDefault(); e.stopPropagation(); zone.classList.add('drag'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      zone.addEventListener(ev, function (e) { e.preventDefault(); e.stopPropagation(); zone.classList.remove('drag'); });
    });
    zone.addEventListener('drop', function (e) {
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) handler(e.dataTransfer.files[0]);
    });
  }
  wireDropzone(els.dropzone, els.fileInput, handleFile);
  wireDropzone(els.oppsDropzone, els.oppsInput, handleOppsFile);
  els.oppsClear.addEventListener('click', clearOpps);

  // ---- Top-right info popups (tutorial / storage / privacy) ----------------
  var popups = [
    { btn: els.helpBtn, pop: els.helpPop },
    { btn: els.cookieBtn, pop: els.cookiePop },
    { btn: els.lockBtn, pop: els.privacyPop }
  ];
  function closeOtherPops(except) {
    popups.forEach(function (p) {
      if (p.pop !== except) { p.pop.hidden = true; p.btn.setAttribute('aria-expanded', 'false'); }
    });
  }
  popups.forEach(function (p) {
    p.btn.addEventListener('click', function (e) {
      e.stopPropagation();
      var show = p.pop.hidden;
      closeOtherPops(p.pop);
      p.pop.hidden = !show;
      p.btn.setAttribute('aria-expanded', String(show));
    });
    var closeBtn = p.pop.querySelector('.pop-close');
    if (closeBtn) closeBtn.addEventListener('click', function () {
      p.pop.hidden = true; p.btn.setAttribute('aria-expanded', 'false');
    });
  });
  document.addEventListener('click', function (e) {
    popups.forEach(function (p) {
      if (p.pop.hidden) return;
      if (!p.pop.contains(e.target) && !p.btn.contains(e.target)) {
        p.pop.hidden = true; p.btn.setAttribute('aria-expanded', 'false');
      }
    });
  });
  function flashCleared() {
    els.cookieCleared.hidden = false;
    setTimeout(function () { els.cookieCleared.hidden = true; }, 2000);
  }
  els.cookieClear.addEventListener('click', function () { clearSavedState(); flashCleared(); });
  els.cookieClearAll.addEventListener('click', function () {
    clearAllStored();       // edits + theme preference
    applyTheme(false);      // reset to light
    flashCleared();
  });

  // ---- Dark mode (preference saved in local storage) -----------------------
  var THEME_KEY = 'rp:theme';
  function applyTheme(dark) {
    document.documentElement.classList.toggle('dark', dark);
    els.themeBtn.textContent = dark ? '☀' : '🌙';
    els.themeBtn.setAttribute('aria-pressed', String(dark));
    els.themeBtn.setAttribute('aria-label', dark ? 'Switch to light mode' : 'Switch to dark mode');
  }
  (function () {
    var saved = null;
    try { saved = localStorage.getItem(THEME_KEY); } catch (e) {}
    applyTheme(saved === 'dark');
  })();
  els.themeBtn.addEventListener('click', function () {
    var dark = !document.documentElement.classList.contains('dark');
    applyTheme(dark);
    try { localStorage.setItem(THEME_KEY, dark ? 'dark' : 'light'); } catch (e) {}
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

  // Per-record eye button jumps straight into edit-actions for that group.
  els.sheetTable.addEventListener('click', function (e) {
    var b = e.target.closest('.row-edit'); if (!b) return;
    openEdit(state.activeSheet, parseInt(b.getAttribute('data-group'), 10));
  });

  // ---- Opportunities popup -------------------------------------------------
  document.addEventListener('click', function (e) {
    var b = e.target.closest('.opps-bubble');
    if (!b || b.classList.contains('zero')) return;
    openOppsPopup(b.getAttribute('data-acct'));
  });
  els.oppsClose.addEventListener('click', function () { els.oppsOverlay.hidden = true; });
  els.oppsOverlay.addEventListener('click', function (e) { if (e.target === els.oppsOverlay) els.oppsOverlay.hidden = true; });
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Escape') return;
    if (!els.oppsOverlay.hidden) els.oppsOverlay.hidden = true;
    else if (!els.tagOverlay.hidden) closeTagPopup();
  });

  function openOppsPopup(acct) {
    var list = (state.opps && state.opps.byAccount[acct]) || [];
    els.oppsTitle.textContent = 'Opportunities for ' + acct + ' (' + list.length + ')';
    var head = '<tr class="head"><td>Opportunity ID</td><td>Opportunity owner</td><td>Stage</td><td>Last modified date</td><td>Owner role</td></tr>';
    var body = list.map(function (o) {
      var link = o.id
        ? '<a href="' + esc(SF + encodeURIComponent(o.id)) + '" target="_blank" rel="noopener noreferrer">' + esc(o.id) + '</a>'
        : '';
      return '<tr><td class="mono">' + link + '</td><td>' + esc(o.owner) + '</td><td>' +
        esc(o.stage) + '</td><td>' + esc(o.modifiedRaw) + '</td><td>' + esc(o.role) + '</td></tr>';
    }).join('');
    els.oppsTable.innerHTML = head + (body || '<tr class="empty"><td colspan="5">No opportunities.</td></tr>');
    els.oppsOverlay.hidden = false;
  }

  // ---- Edit-mode controls --------------------------------------------------
  els.editClose.addEventListener('click', closeEdit);
  els.prevGroup.addEventListener('click', function () { navGroup(-1); });
  els.nextGroup.addEventListener('click', function () { navGroup(1); });
  els.remarksInput.addEventListener('input', function () { currentGroup().remarks = els.remarksInput.value; saveState(); });
  els.actionButtons.addEventListener('click', function (e) {
    var b = e.target.closest('.action-btn'); if (!b) return;
    setAction(b.getAttribute('data-action'));
  });
  els.editTable.addEventListener('change', function (e) {
    var r = e.target.closest('input[type="radio"]'); if (!r) return;
    setPrimary(parseInt(r.getAttribute('data-row'), 10));
  });
  els.editTable.addEventListener('click', function (e) {
    var d = e.target.closest('.row-del'); if (!d) return;
    deleteRecord(parseInt(d.getAttribute('data-row'), 10));
  });
  els.editUndo.addEventListener('click', undoDelete);
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
    var t0 = now();
    state.result = RP.buildReport(RP.parseCSV(text));
    state.activeSheet = 0;
    state.reviewSheet = state.result.review;
    lastName = deriveName(name);
    state.csvKey = STORAGE_PREFIX + hashText(text);
    // Per-group review fields + a stable uid; reset defaults before restore.
    // reviewInclude selects which report sections (sheets) feed the Review tab.
    state.reviewInclude = state.result.sheets.map(function () { return true; });
    state.reviewSearch = {};
    state.tagFilter = {};
    state.tagExclude = {};
    state.tagPool = [];
    eachGroup(function (g, si) {
      g.uid = state.result.sheets[si].key + ':' + g.key;
      g.sheetIndex = si;
      g.peopleTags = [];
      g.nextStep = 'None';
      g.deletedRows = [];
    });
    if (state.opps) applyOpps();  // re-attach opportunities to the fresh model
    restoreState();               // re-apply any saved edits for this exact file
    ensurePresets();              // Default preset for every target if none saved
    state.processMs = Math.round(now() - t0);
    renderStats(state.result.stats);
    renderWarnings(state.result.warnings);
    buildTabs();
    selectSheet(0);
    els.dashCard.hidden = false;
    els.viewer.hidden = false;
    els.viewer.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  // Visit every group across all sheets: fn(group, sheetIndex, groupIndex).
  function eachGroup(fn) {
    if (!state.result) return;
    state.result.sheets.forEach(function (s, si) {
      s.groups.forEach(function (g, gi) { fn(g, si, gi); });
    });
  }
  function allGroups() { var out = []; eachGroup(function (g) { out.push(g); }); return out; }

  // ---- Opportunities: parse, index, attach --------------------------------
  var OPP = { id: 0, owner: 2, role: 3, account: 5, stage: 13, modified: 23 }; // cols A,C,D,F,N,X

  function handleOppsFile(file) {
    var reader = new FileReader();
    reader.onerror = function () { showError('Could not read the Opportunities file.'); };
    reader.onload = function () {
      try {
        var rows = RP.parseCSV(String(reader.result));
        state.opps = indexOpps(rows);
        state.oppsName = file.name;
        els.oppsName.hidden = false;
        els.oppsName.innerHTML = 'Loaded: <strong>' + esc(file.name) + '</strong>'; // the linked-count stat lives on the dashboard
        els.oppsClear.hidden = false;
        if (state.result) { applyOpps(); renderStats(state.result.stats); rerenderActive(); saveState(); }
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
      if (state.reviewSheet) removeOppsField(state.reviewSheet);
      state.result.sheets.forEach(function (s) {
        s.groups.forEach(function (g) { g.rows.forEach(function (r) { r.opps = null; r.oppsSummary = ''; }); });
      });
      renderStats(state.result.stats);
      rerenderActive();
      saveState();
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
        stage: String(r[OPP.stage] == null ? '' : r[OPP.stage]).trim(),
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
    if (state.reviewSheet) ensureOppsField(state.reviewSheet); // review shares the same rows
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
    if (state.edit.active) renderEditGroup();
    else if (isReviewActive()) renderReview();
    else renderSheet();
  }

  // ---- Local persistence ---------------------------------------------------
  // Edits are remembered in the browser's own localStorage (NOT cookies, so
  // nothing is ever transmitted), keyed to a fingerprint of the uploaded CSV.
  // Re-uploading the same file after a reload restores actions, primary
  // choices, removed records and the customize layout.
  var STORAGE_PREFIX = 'rp:v1:';

  function hashText(t) {
    t = String(t == null ? '' : t);
    var h = 5381;
    for (var i = 0; i < t.length; i++) { h = ((h << 5) + h + t.charCodeAt(i)) | 0; }
    return (h >>> 0).toString(16) + '.' + t.length;
  }

  function rowAcct(row) {
    return row && row.raw ? String(row.raw[RP.COL.accountId] == null ? '' : row.raw[RP.COL.accountId]).trim() : '';
  }

  function layoutOf(sheet) {
    return { fields: sheet.fields.map(function (f) { return { id: f.id, label: f.label }; }), displayCount: sheet.displayCount };
  }

  function saveState() {
    if (!state.result || !state.csvKey) return;
    try {
      var data = { sheets: {}, tagPool: state.tagPool.slice() };
      state.result.sheets.forEach(function (sheet) {
        var groups = {};
        sheet.groups.forEach(function (g) {
          var primary = g.rows.filter(function (r) { return r.isPrimary; })[0];
          groups[g.key] = {
            order: g.rows.map(rowAcct),
            primaryId: primary ? rowAcct(primary) : null,
            action: g.action,
            actionChosen: !!g.actionChosen,
            remarks: g.remarks || '',
            peopleTags: (g.peopleTags || []).slice(),
            nextStep: g.nextStep || 'None'
          };
        });
        var lay = layoutOf(sheet);
        data.sheets[sheet.key] = {
          fields: lay.fields, displayCount: lay.displayCount,
          presets: sheet.presets, activePreset: sheet.activePreset,
          groups: groups
        };
      });
      if (state.reviewSheet) {
        var rlay = layoutOf(state.reviewSheet);
        data.review = {
          fields: rlay.fields, displayCount: rlay.displayCount,
          presets: state.reviewSheet.presets, activePreset: state.reviewSheet.activePreset,
          include: state.reviewInclude.slice()
        };
      }
      localStorage.setItem(state.csvKey, JSON.stringify(data));
    } catch (e) { /* storage unavailable/full — silently skip */ }
  }

  function restoreState() {
    if (!state.result || !state.csvKey) return;
    var data;
    try {
      var raw = localStorage.getItem(state.csvKey);
      if (!raw) return;
      data = JSON.parse(raw);
    } catch (e) { return; }
    if (!data || !data.sheets) return;

    if (Array.isArray(data.tagPool)) state.tagPool = data.tagPool.slice();

    state.result.sheets.forEach(function (sheet) {
      var s = data.sheets[sheet.key];
      if (!s) return;
      restoreFields(sheet, s);
      restorePresets(sheet, s);
      sheet.groups.forEach(function (g) {
        var gs = s.groups && s.groups[g.key];
        if (!gs) return;
        // Rebuild rows in the saved order; account IDs absent from the saved
        // order were removed by the user. They drop out of g.rows but are kept
        // in g.deletedRows so the removal can still be undone after a reload.
        if (Array.isArray(gs.order) && gs.order.length) {
          var byAcct = {};
          g.rows.forEach(function (r) { byAcct[rowAcct(r)] = r; });
          var newRows = gs.order.map(function (id) { return byAcct[id]; }).filter(Boolean);
          if (newRows.length) {
            var kept = {};
            newRows.forEach(function (r) { kept[rowAcct(r)] = 1; });
            g.deletedRows = g.rows.filter(function (r) { return !kept[rowAcct(r)]; });
            g.rows = newRows;
          }
        }
        g.rows.forEach(function (r) {
          var isP = gs.primaryId != null && rowAcct(r) === gs.primaryId;
          if (gs.primaryId != null) { r.isPrimary = isP; r.classification = isP ? 'Primary' : 'Duplicate'; }
        });
        if (typeof gs.action === 'string') g.action = gs.action;
        g.actionChosen = !!gs.actionChosen;
        if (typeof gs.remarks === 'string') g.remarks = gs.remarks;
        if (Array.isArray(gs.peopleTags)) g.peopleTags = gs.peopleTags.slice();
        if (typeof gs.nextStep === 'string') g.nextStep = gs.nextStep;
      });
    });

    if (data.review && state.reviewSheet) {
      restoreFields(state.reviewSheet, data.review);
      restorePresets(state.reviewSheet, data.review);
      if (Array.isArray(data.review.include) && data.review.include.length === state.reviewInclude.length) {
        state.reviewInclude = data.review.include.map(function (v) { return !!v; });
      }
    }
  }

  function restorePresets(sheet, s) {
    if (s && Array.isArray(s.presets) && s.presets.length) {
      sheet.presets = s.presets;
      sheet.activePreset = (typeof s.activePreset === 'number' &&
        s.activePreset >= 0 && s.activePreset < s.presets.length) ? s.activePreset : 0;
    }
  }

  // ---- Column presets (per customization target) ---------------------------
  function presetTargets() {
    var t = state.result.sheets.slice();
    if (state.reviewSheet) t.push(state.reviewSheet);
    return t;
  }
  function ensurePresets() {
    presetTargets().forEach(function (sheet) {
      if (!Array.isArray(sheet.presets) || !sheet.presets.length) {
        sheet.presets = [{ name: 'Default', layout: layoutOf(sheet) }];
        sheet.activePreset = 0;
      }
    });
  }
  function layoutSig(lay) {
    return lay.displayCount + '|' + lay.fields.map(function (f) { return f.id + '=' + f.label; }).join(',');
  }
  function presetIsDirty(sheet) {
    if (!sheet.presets || !sheet.presets.length) return false;
    var active = sheet.presets[sheet.activePreset] || sheet.presets[0];
    return layoutSig(layoutOf(sheet)) !== layoutSig(active.layout);
  }
  function applyPreset(sheet, idx) {
    var p = sheet.presets[idx];
    if (!p) return;
    restoreFields(sheet, p.layout);   // reorders / relabels existing fields
    sheet.activePreset = idx;
  }
  function saveNewPreset(sheet) {
    var n = sheet.presets.filter(function (p) { return p.name !== 'Default'; }).length + 1;
    sheet.presets.push({ name: 'Preset ' + n, layout: layoutOf(sheet) });
    sheet.activePreset = sheet.presets.length - 1;
  }

  function restoreFields(sheet, s) {
    if (!s || !Array.isArray(s.fields)) return;
    var byId = {};
    sheet.fields.forEach(function (f) { byId[f.id] = f; });
    var savedCut = typeof s.displayCount === 'number' ? s.displayCount : s.fields.length;
    var ordered = [], newDisplay = 0;
    s.fields.forEach(function (sf, i) {
      var f = byId[sf.id];
      if (f && !f._seen) {
        if (typeof sf.label === 'string' && sf.label) f.label = sf.label;
        ordered.push(f); f._seen = true;
        if (i < savedCut) newDisplay++; // count only saved-displayed fields that still exist
      }
    });
    sheet.fields.forEach(function (f) { if (!f._seen) ordered.push(f); }); // new fields keep their place at the end (hidden)
    ordered.forEach(function (f) { delete f._seen; });
    if (ordered.length === sheet.fields.length) {
      sheet.fields = ordered;
      sheet.displayCount = Math.max(1, Math.min(newDisplay || savedCut, sheet.fields.length));
    }
  }

  function clearKeys(prefix) {
    try {
      var keys = [];
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (k && k.indexOf(prefix) === 0) keys.push(k);
      }
      keys.forEach(function (k) { localStorage.removeItem(k); });
    } catch (e) { /* ignore */ }
  }
  function clearSavedState() { clearKeys(STORAGE_PREFIX); }      // per-file edits only
  function clearAllStored() { clearKeys('rp:'); }               // edits + theme + anything else

  function renderStats(s) {
    var items = [
      { cls: 's1', num: s.section1Groups, lbl: 'Likely duplicates' },
      { cls: 's2', num: s.section2Groups, lbl: 'Problematic duplicates' },
      { cls: 's3', num: s.unclassifiedGroups, lbl: 'Unclassified' },
      { cls: '', num: s.totalGroups, lbl: 'Account groups total' },
      { cls: '', num: s.totalDataRows, lbl: 'Records read' }
    ];
    if (state.opps) {
      var accts = Object.keys(state.opps.byAccount).length;
      items.push({ cls: 's4', num: state.opps.count, lbl: 'Opportunities · ' + accts + ' account' + (accts === 1 ? '' : 's') });
    }
    if (state.processMs != null) {
      items.push({ cls: 's5', num: state.processMs, lbl: 'Milliseconds to process' });
    }
    els.stats.innerHTML = items.map(function (it) {
      return '<div class="stat ' + it.cls + '"><div class="num">' + it.num + '</div><div class="lbl">' + esc(it.lbl) + '</div></div>';
    }).join('');
  }

  function renderWarnings(warnings) {
    if (!warnings || !warnings.length) { els.warnings.hidden = true; return; }
    els.warnings.hidden = false;
    els.warningList.innerHTML = warnings.map(function (w) { return '<li>' + esc(w) + '</li>'; }).join('');
  }

  function reviewIndex() { return state.result.sheets.length; }
  function isReviewActive() { return state.activeSheet === reviewIndex(); }

  function buildTabs() {
    var tabs = state.result.sheets.map(function (s, i) {
      return '<button type="button" role="tab" class="tab" data-index="' + i + '">' +
        esc(s.title) + ' <span class="tab-count">' + s.groups.length + '</span></button>';
    });
    tabs.push('<button type="button" role="tab" class="tab tab-review" data-index="' + reviewIndex() + '">' +
      'Review <span class="tab-count">' + state.result.stats.totalGroups + '</span></button>');
    els.sheetTabs.innerHTML = tabs.join('');
  }

  function selectSheet(i) {
    state.activeSheet = i;
    Array.prototype.forEach.call(els.sheetTabs.children, function (tab, idx) {
      var on = idx === i;
      tab.classList.toggle('active', on);
      tab.setAttribute('aria-selected', String(on));
    });
    if (isReviewActive()) {
      els.sheetScroll.hidden = true;
      els.reviewView.hidden = false;
      els.viewer.classList.add('review-mode'); // results flow down the whole page
      els.editBtn.hidden = true;
      els.customizeBtn.hidden = false;
      renderReview();
      return;
    }
    els.reviewView.hidden = true;
    els.sheetScroll.hidden = false;
    els.viewer.classList.remove('review-mode');
    els.customizeBtn.hidden = false;
    var sheet = state.result.sheets[i];
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
    // A leading control column (eye button) is shown on-screen for editable
    // sheets only — it is never part of the field model, view export, or edit
    // table.
    var withEye = sheet.editable && sheet.groups.length > 0;
    var span = Math.max(cols.length, 1) + (withEye ? 1 : 0);
    var html = '';
    html += '<tr class="title"><td colspan="' + span + '">' + esc(sheet.title) + '</td></tr>';
    html += '<tr class="desc"><td colspan="' + span + '">' + esc(sheet.description) + '</td></tr>';
    html += '<tr class="head">' + (withEye ? '<td class="pv-eye-head"></td>' : '') +
      cols.map(function (c) { return '<td>' + esc(c.label) + '</td>'; }).join('') + '</tr>';
    if (!sheet.groups.length) {
      html += '<tr class="empty"><td colspan="' + span + '">No records in this section.</td></tr>';
    } else {
      sheet.groups.forEach(function (g, gi) {
        if (gi > 0) html += '<tr class="blank"><td colspan="' + span + '"></td></tr>';
        g.rows.forEach(function (row, ri) {
          // One eye button per group, on the group's first row only.
          var eye = withEye
            ? '<td class="pv-eye">' + (ri === 0
                ? '<button type="button" class="row-edit" data-group="' + gi +
                  '" aria-label="Open this group in edit actions" title="Edit this group">&#128065;</button>'
                : '') + '</td>'
            : '';
          html += '<tr' + (row.isPrimary ? ' class="primary"' : '') + '>' + eye +
            cols.map(function (c) { return td(c, g, row, ri); }).join('') + '</tr>';
        });
      });
    }
    els.sheetTable.className = 'pv theme-' + sheet.theme.name;
    els.sheetTable.innerHTML = html;
  }

  // ---- Review tab ----------------------------------------------------------
  // Pools every group and partitions it by the assigned action + remarks:
  //   0 Merge/no remarks · 1 Merge/remarks · 2 Evaluate or None-with-remarks
  //   3 Close · 4 None (no remarks)
  function reviewPartition(g) {
    var a = g.action || 'None';
    var hasRemarks = !!(g.remarks && g.remarks.trim());
    if (a === 'Merge') return hasRemarks ? 1 : 0;
    if (a === 'Evaluate' || (a === 'None' && hasRemarks)) return 2;
    if (a === 'Close') return 3;
    return 4; // None, no remarks
  }

  function reviewDataFields() { return RP.displayedFields(state.reviewSheet); }

  function groupMatchesColumn(g, f, textLower) {
    for (var i = 0; i < g.rows.length; i++) {
      var v = RP.fieldValue(f, g, g.rows[i], i);
      if (v && String(v).toLowerCase().indexOf(textLower) >= 0) return true;
    }
    return false;
  }
  // Does the column hold any value for this group (across all its rows)?
  function columnHasValue(g, f) {
    for (var i = 0; i < g.rows.length; i++) {
      var v = RP.fieldValue(f, g, g.rows[i], i);
      if (v != null && String(v).trim() !== '') return true;
    }
    return false;
  }
  function matchesSearch(g) {
    var fields = reviewDataFields();
    for (var fid in state.reviewSearch) {
      if (!Object.prototype.hasOwnProperty.call(state.reviewSearch, fid)) continue;
      var raw = (state.reviewSearch[fid] || '').trim();
      if (!raw) continue;
      var f = fields.filter(function (x) { return x.id === fid; })[0];
      if (!f) continue; // column no longer displayed → ignore its filter
      // Special tokens: "*" = any non-empty value, "-" = only empty. (The
      // leading backslash form is accepted too, since some editors escape them.)
      if (raw === '*' || raw === '\\*') { if (!columnHasValue(g, f)) return false; continue; }
      if (raw === '-' || raw === '\\-') { if (columnHasValue(g, f)) return false; continue; }
      if (!groupMatchesColumn(g, f, raw.toLowerCase())) return false;
    }
    return true;
  }
  function matchesTagFilter(g) {
    var tags = g.peopleTags || [];
    // Exclusions: drop any group that carries an excluded tag.
    for (var ex in state.tagExclude) {
      if (state.tagExclude[ex] && tags.indexOf(ex) >= 0) return false;
    }
    var sel = Object.keys(state.tagFilter).filter(function (k) { return state.tagFilter[k]; });
    var any = sel.indexOf(ANY_TAG) >= 0;
    var real = sel.filter(function (t) { return t !== ANY_TAG; });
    if (!real.length) return true; // nothing (or "Any" alone) selected → all groups
    for (var i = 0; i < real.length; i++) { if (tags.indexOf(real[i]) < 0) return false; }
    return any ? true : tags.length === real.length; // Any = superset; else exact set
  }
  // Tag chips cycle off → include → exclude → off (the "Any" wildcard only
  // toggles include, since excluding "any tag" is meaningless).
  function cycleTagFilter(tag) {
    if (tag === ANY_TAG) { state.tagFilter[tag] = !state.tagFilter[tag]; return; }
    if (state.tagFilter[tag]) { state.tagFilter[tag] = false; state.tagExclude[tag] = true; }
    else if (state.tagExclude[tag]) { state.tagExclude[tag] = false; }
    else { state.tagFilter[tag] = true; }
  }
  // A group is on screen when its report section is included and it passes the
  // column search + tag filter. This is the target of "apply to all shown" and
  // "copy results". The five partition tables below just re-group these.
  function sectionIncluded(g) { return state.reviewInclude[g.sheetIndex] !== false; }
  function visibleGroups() {
    return allGroups().filter(function (g) {
      return sectionIncluded(g) && matchesSearch(g) && matchesTagFilter(g);
    });
  }
  function partitionBuckets() {
    var buckets = [[], [], [], [], []];
    visibleGroups().forEach(function (g) { buckets[reviewPartition(g)].push(g); });
    return buckets;
  }
  function findByUid(uid) {
    var found = null;
    eachGroup(function (g, si, gi) { if (g.uid === uid) found = { g: g, si: si, gi: gi }; });
    return found;
  }

  function renderReview() {
    renderReviewToolbar();
    renderReviewSearchRow();
    renderReviewResults();
  }

  function renderReviewToolbar() {
    // Include/exclude by report SECTION (Likely / Problematic / Unclassified).
    var incHtml = state.result.sheets.map(function (s, i) {
      var on = state.reviewInclude[i] !== false;
      return '<button type="button" class="rv-chip rv-inc' + (on ? ' active' : '') +
        '" data-inc="' + i + '" aria-pressed="' + (on ? 'true' : 'false') + '">' +
        esc(s.title) + '</button>';
    }).join('');

    var filterTags = [ANY_TAG].concat(state.tagPool);
    var tfHtml = filterTags.map(function (t) {
      var inc = !!state.tagFilter[t];
      var exc = !!state.tagExclude[t];
      var cls = 'rv-chip rv-tf' + (t === ANY_TAG ? ' rv-any' : '') + (inc ? ' active' : '') + (exc ? ' rv-ex' : '');
      var pressed = inc ? 'true' : (exc ? 'mixed' : 'false');
      return '<button type="button" class="' + cls + '" data-tf="' + esc(t) + '" aria-pressed="' + pressed +
        '">' + (exc ? '✕ ' : '') + esc(t) + '</button>';
    }).join('');
    // Non-"Any" chips cycle include → exclude → off, so add a hint.
    var tfHint = state.tagPool.length ? '<span class="rv-hint">click a tag: include → exclude → off</span>' : '';

    var nextHtml = '<select id="rvNextAll" class="rv-select" aria-label="Set next step for all shown groups">' +
      '<option value="">Next step for all…</option>' +
      NEXT_STEPS.map(function (s) { return '<option value="' + esc(s) + '">' + esc(s) + '</option>'; }).join('') +
      '</select>';

    els.reviewToolbar.innerHTML =
      '<div class="rv-block"><span class="rv-block-label">Sections</span><div class="rv-chips">' + incHtml + '</div></div>' +
      '<div class="rv-block"><span class="rv-block-label">Filter tags</span><div class="rv-chips">' + tfHtml + '</div>' + tfHint + '</div>' +
      '<div class="rv-block"><span class="rv-block-label">Apply to all shown</span><div class="rv-chips">' +
        '<button type="button" id="rvTagAll" class="btn">🏷 Tag all shown…</button>' + nextHtml +
      '</div></div>';
  }

  function renderReviewSearchRow() {
    var fields = reviewDataFields();
    els.reviewSearch.innerHTML = '<span class="rv-block-label">Search columns ' +
      '<span class="rv-hint">— <code>*</code> = any value, <code>-</code> = empty</span></span>' +
      '<div class="rv-search-grid">' + fields.map(function (f) {
        return '<label class="rv-search-col"><span>' + esc(f.label) + '</span>' +
          '<input type="search" class="rv-search-input" data-fid="' + esc(f.id) + '" value="' +
          esc(state.reviewSearch[f.id] || '') + '" placeholder="filter…" autocomplete="off" /></label>';
      }).join('') + '</div>';
  }

  function reviewGroupRowsHtml(g, si, gi, fields, withEye) {
    var html = '';
    g.rows.forEach(function (row, ri) {
      var lead = withEye
        ? '<td class="pv-eye">' + (ri === 0
            ? '<button type="button" class="row-edit" data-uid="' + esc(g.uid) +
              '" aria-label="Open this group in edit actions" title="Edit this group">&#128065;</button>'
            : '') + '</td>'
        : '';
      var tagCell, nextCell;
      if (ri === 0) {
        var tags = g.peopleTags || [];
        var tagLabel = tags.length
          ? tags.map(function (t) { return '<span class="tag-pill">' + esc(t) + '</span>'; }).join(' ')
          : '<span class="tag-add-hint">＋ Add tags</span>';
        tagCell = '<td class="rv-tags"><button type="button" class="tag-btn" data-uid="' + esc(g.uid) +
          '" title="Assign people tags">' + tagLabel + '</button></td>';
        nextCell = '<td class="rv-next"><select class="next-step" data-uid="' + esc(g.uid) + '" aria-label="Next step">' +
          NEXT_STEPS.map(function (s) {
            return '<option value="' + esc(s) + '"' + ((g.nextStep || 'None') === s ? ' selected' : '') + '>' + esc(s) + '</option>';
          }).join('') + '</select></td>';
      } else {
        tagCell = '<td class="rv-tags"></td>';
        nextCell = '<td class="rv-next"></td>';
      }
      html += '<tr' + (row.isPrimary ? ' class="primary"' : '') + '>' + lead +
        fields.map(function (c) { return td(c, g, row, ri); }).join('') + tagCell + nextCell + '</tr>';
    });
    return html;
  }

  function renderReviewResults() {
    var fields = reviewDataFields();
    var buckets = partitionBuckets();
    var loc = {};
    eachGroup(function (g, si, gi) { loc[g.uid] = { si: si, gi: gi }; });
    var withEye = true;
    var span = 1 + fields.length + 2; // eye + data + (tags, next)

    var shownCount = visibleGroups().length;
    var head = '<div class="rv-results-head">' +
      '<span class="rv-count">Showing <strong>' + shownCount + '</strong> group' + (shownCount === 1 ? '' : 's') + '</span>' +
      '<button type="button" class="btn review-copy" data-copy="all">⧉ Copy results</button></div>';

    var tables = '';
    REVIEW_TABLES.forEach(function (t, i) {
      var groups = buckets[i];
      var block = '<div class="review-table-block">' +
        '<div class="review-table-head"><h3>' + esc(t.title) +
        ' <span class="tab-count">' + groups.length + '</span></h3>' +
        '<button type="button" class="btn review-copy" data-copy="' + i + '"' + (groups.length ? '' : ' disabled') + '>⧉ Copy</button></div>';
      var rows = '<tr class="head"><td class="pv-eye-head"></td>' +
        fields.map(function (c) { return '<td>' + esc(c.label) + '</td>'; }).join('') +
        '<td>People tags</td><td>Next steps</td></tr>';
      if (!groups.length) {
        rows += '<tr class="empty"><td colspan="' + span + '">No groups.</td></tr>';
      } else {
        groups.forEach(function (g, gj) {
          if (gj > 0) rows += '<tr class="blank"><td colspan="' + span + '"></td></tr>';
          var l = loc[g.uid] || {};
          rows += reviewGroupRowsHtml(g, l.si, l.gi, fields, withEye);
        });
      }
      block += '<div class="table-scroll"><table class="pv theme-gray review-table">' + rows + '</table></div></div>';
      tables += block;
    });

    els.reviewResults.innerHTML = head + (tables || '<p class="muted">No tables selected.</p>');
  }

  // ---- Review: clipboard (email-friendly) ----------------------------------
  function reviewCopyMatrix(groups) {
    var fields = reviewDataFields();
    var header = fields.map(function (f) { return f.label; });
    var body = [];
    groups.forEach(function (g) {
      g.rows.forEach(function (row, ri) {
        body.push({ primary: !!row.isPrimary, first: ri === 0, cells: fields.map(function (f) { return String(RP.fieldValue(f, g, row, ri)); }) });
      });
    });
    return { header: header, body: body };
  }
  // Email-friendly HTML. Uses presentational tags/attributes (no inline
  // `style=`) so it stays within the page's strict style-src CSP, while still
  // pasting as a bordered table into Outlook / Gmail.
  function copyHtml(m, caption) {
    var cols = m.header.length;
    var th = m.header.map(function (h) { return '<th align="left">' + esc(h) + '</th>'; }).join('');
    var trs = m.body.map(function (r, i) {
      var sep = (r.first && i > 0) ? '<tr><td colspan="' + cols + '">&nbsp;</td></tr>' : '';
      var tds = r.cells.map(function (c) { return '<td>' + (r.primary ? '<b>' + esc(c) + '</b>' : esc(c)) + '</td>'; }).join('');
      return sep + '<tr>' + tds + '</tr>';
    }).join('');
    return (caption ? '<p><b>' + esc(caption) + '</b></p>' : '') +
      '<table border="1" cellspacing="0" cellpadding="4"><thead><tr>' + th + '</tr></thead><tbody>' + trs + '</tbody></table>';
  }
  function copyText(m, caption) {
    var lines = [];
    if (caption) lines.push(caption);
    lines.push(m.header.join('\t'));
    m.body.forEach(function (r, i) {
      if (r.first && i > 0) lines.push(''); // blank line between groups
      lines.push(r.cells.join('\t'));
    });
    return lines.join('\n');
  }
  function legacyCopy(text) {
    return new Promise(function (resolve) {
      try {
        var ta = document.createElement('textarea');
        ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
        document.body.appendChild(ta); ta.focus(); ta.select();
        document.execCommand('copy'); document.body.removeChild(ta);
      } catch (e) { /* ignore */ }
      resolve();
    });
  }
  function writeClipboard(html, text) {
    if (navigator.clipboard && typeof window.ClipboardItem === 'function') {
      try {
        var item = new window.ClipboardItem({
          'text/html': new Blob([html], { type: 'text/html' }),
          'text/plain': new Blob([text], { type: 'text/plain' })
        });
        return navigator.clipboard.write([item]).catch(function () { return legacyCopy(text); });
      } catch (e) { return legacyCopy(text); }
    }
    return legacyCopy(text);
  }
  function flashCopied(btn) {
    if (!btn) return;
    var prev = btn.textContent;
    btn.textContent = 'Copied ✓';
    btn.classList.add('copied');
    setTimeout(function () { btn.textContent = prev; btn.classList.remove('copied'); }, 1500);
  }
  function doCopy(which, btn) {
    var groups, caption;
    if (which === 'all') { groups = visibleGroups(); caption = 'Review — ' + groups.length + ' group(s)'; }
    else {
      var i = parseInt(which, 10);
      groups = partitionBuckets()[i] || [];
      caption = REVIEW_TABLES[i].title + ' — ' + groups.length + ' group(s)';
    }
    if (!groups.length) return;
    var m = reviewCopyMatrix(groups);
    writeClipboard(copyHtml(m, caption), copyText(m, caption)).then(function () { flashCopied(btn); });
  }

  // ---- Review: event wiring ------------------------------------------------
  els.reviewToolbar.addEventListener('click', function (e) {
    var inc = e.target.closest('.rv-inc');
    if (inc) { var i = parseInt(inc.getAttribute('data-inc'), 10); state.reviewInclude[i] = !state.reviewInclude[i]; saveState(); renderReview(); return; }
    var tf = e.target.closest('.rv-tf');
    if (tf) { cycleTagFilter(tf.getAttribute('data-tf')); renderReview(); return; }
    if (e.target.closest('#rvTagAll')) { openTagPopup('*'); return; }
  });
  els.reviewToolbar.addEventListener('change', function (e) {
    var sel = e.target.closest('#rvNextAll'); if (!sel) return;
    var v = sel.value; if (NEXT_STEPS.indexOf(v) < 0) return;
    visibleGroups().forEach(function (g) { g.nextStep = v; });
    saveState(); renderReview();
  });
  els.reviewSearch.addEventListener('input', function (e) {
    var inp = e.target.closest('.rv-search-input'); if (!inp) return;
    state.reviewSearch[inp.getAttribute('data-fid')] = inp.value;
    renderReviewResults(); // only the results, so the focused input keeps focus
  });
  els.reviewResults.addEventListener('click', function (e) {
    var copy = e.target.closest('.review-copy');
    if (copy) { doCopy(copy.getAttribute('data-copy'), copy); return; }
    var tagB = e.target.closest('.tag-btn');
    if (tagB) { var f = findByUid(tagB.getAttribute('data-uid')); if (f) openTagPopup(f.g); return; }
    var eye = e.target.closest('.row-edit');
    if (eye) { var g2 = findByUid(eye.getAttribute('data-uid')); if (g2) openEdit(g2.si, g2.gi); return; }
  });
  els.reviewResults.addEventListener('change', function (e) {
    var sel = e.target.closest('.next-step'); if (!sel) return;
    var f = findByUid(sel.getAttribute('data-uid')); if (!f) return;
    f.g.nextStep = sel.value; saveState();
  });

  // ---- Review: people-tags popup -------------------------------------------
  function openTagPopup(target) {
    state.tagTarget = target;          // a group object, or '*' for bulk-apply
    state.tagBulkPick = {};
    renderTagPopup();
    els.tagOverlay.hidden = false;
    els.tagInput.value = '';
    els.tagInput.focus();
  }
  function closeTagPopup() {
    els.tagOverlay.hidden = true;
    state.tagTarget = null;
    if (isReviewActive()) renderReview(); // new pool tags appear as filter chips
  }

  function renderTagPopup() {
    var bulk = state.tagTarget === '*';
    if (bulk) {
      var n = visibleGroups().length;
      els.tagTitle.textContent = 'Tag all shown groups';
      els.tagHint.textContent = 'Pick or create tags, then apply them to all ' + n + ' shown group(s).';
      els.tagCurrent.hidden = true;
    } else {
      var g = state.tagTarget;
      els.tagTitle.textContent = 'People tags';
      els.tagHint.textContent = 'Assign, create, or remove tags for this group.';
      els.tagCurrent.hidden = false;
      var tags = g.peopleTags || [];
      els.tagCurrent.innerHTML = tags.length
        ? tags.map(function (t) {
            return '<span class="tag-pill removable" data-rm="' + esc(t) + '">' + esc(t) +
              '<button type="button" class="tag-rm" data-rm="' + esc(t) + '" aria-label="Remove ' + esc(t) + '">×</button></span>';
          }).join(' ')
        : '<span class="muted">No tags yet.</span>';
    }
    els.tagPool.innerHTML = state.tagPool.length
      ? state.tagPool.map(function (t) {
          var on = bulk ? !!state.tagBulkPick[t] : (state.tagTarget.peopleTags || []).indexOf(t) >= 0;
          return '<button type="button" class="rv-chip tag-pool-chip' + (on ? ' active' : '') + '" data-tag="' + esc(t) + '">' + esc(t) + '</button>';
        }).join('')
      : '<span class="muted">No tags created yet — type one below.</span>';
    els.tagApply.hidden = !bulk;
    if (bulk) {
      var picks = Object.keys(state.tagBulkPick).filter(function (k) { return state.tagBulkPick[k]; }).length;
      els.tagApply.innerHTML = '<button type="button" id="tagApplyBtn" class="btn primary"' + (picks ? '' : ' disabled') +
        '>Apply ' + picks + ' tag(s) to all shown</button>';
    }
  }

  function poolAdd(tag) { if (state.tagPool.indexOf(tag) < 0) { state.tagPool.push(tag); } }

  function addTagFromInput() {
    var v = els.tagInput.value.trim();
    if (!v) return;
    poolAdd(v);
    if (state.tagTarget === '*') { state.tagBulkPick[v] = true; }
    else {
      var g = state.tagTarget;
      if (!g.peopleTags) g.peopleTags = [];
      if (g.peopleTags.indexOf(v) < 0) g.peopleTags.push(v);
      saveState();
    }
    els.tagInput.value = '';
    els.tagInput.focus();
    renderTagPopup();
    if (state.tagTarget !== '*') renderReviewResults();
    if (state.tagTarget === '*') saveState(); // persist the enlarged pool
  }
  function togglePoolTag(tag) {
    if (state.tagTarget === '*') { state.tagBulkPick[tag] = !state.tagBulkPick[tag]; renderTagPopup(); return; }
    var g = state.tagTarget;
    if (!g.peopleTags) g.peopleTags = [];
    var i = g.peopleTags.indexOf(tag);
    if (i >= 0) g.peopleTags.splice(i, 1); else g.peopleTags.push(tag);
    saveState(); renderTagPopup(); renderReviewResults();
  }
  function removeTagFromGroup(tag) {
    var g = state.tagTarget; if (g === '*' || !g.peopleTags) return;
    var i = g.peopleTags.indexOf(tag);
    if (i >= 0) { g.peopleTags.splice(i, 1); saveState(); renderTagPopup(); renderReviewResults(); }
  }
  function applyBulkTags() {
    var picks = Object.keys(state.tagBulkPick).filter(function (k) { return state.tagBulkPick[k]; });
    if (!picks.length) return;
    visibleGroups().forEach(function (g) {
      if (!g.peopleTags) g.peopleTags = [];
      picks.forEach(function (t) { if (g.peopleTags.indexOf(t) < 0) g.peopleTags.push(t); });
    });
    saveState();
    closeTagPopup();
    renderReview();
  }

  els.tagAdd.addEventListener('click', addTagFromInput);
  els.tagInput.addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); addTagFromInput(); } });
  els.tagClose.addEventListener('click', closeTagPopup);
  els.tagOverlay.addEventListener('click', function (e) {
    if (e.target === els.tagOverlay) { closeTagPopup(); return; }
    var pool = e.target.closest('.tag-pool-chip');
    if (pool) { togglePoolTag(pool.getAttribute('data-tag')); return; }
    var rm = e.target.closest('.tag-rm');
    if (rm) { removeTagFromGroup(rm.getAttribute('data-rm')); return; }
    if (e.target.closest('#tagApplyBtn')) { applyBulkTags(); return; }
  });

  // ---- Edit-actions mode ---------------------------------------------------
  function currentSheet() { return state.result.sheets[state.edit.sheet]; }
  function currentGroup() { return currentSheet().groups[state.edit.group]; }

  function openEdit(sheetIndex, groupIndex) {
    var sheet = state.result.sheets[sheetIndex];
    if (!sheet.editable || !sheet.groups.length) return;
    var gi = groupIndex || 0;
    if (gi < 0) gi = 0;
    if (gi >= sheet.groups.length) gi = sheet.groups.length - 1;
    state.edit = { active: true, sheet: sheetIndex, group: gi };
    els.editOverlay.hidden = false;
    renderEditGroup();
  }

  function closeEdit() {
    // Leaving the group: apply the primary-to-top reorder offscreen.
    RP.reorderPrimaryTop(currentGroup());
    saveState();
    state.edit.active = false;
    els.editOverlay.hidden = true;
    // Edit can be opened from any tab (incl. Review); re-render the active one.
    if (isReviewActive()) renderReview(); else renderSheet();
  }

  function navGroup(delta) {
    var sheet = currentSheet();
    var next = state.edit.group + delta;
    if (next < 0 || next >= sheet.groups.length) return;
    // Reorder the group we're leaving offscreen, so its rows show the primary
    // on top next time it's viewed — but never while it's on screen.
    RP.reorderPrimaryTop(currentGroup());
    saveState();
    state.edit.group = next;
    renderEditGroup();
  }

  // An Unclassified group needs an action before advancing, and a primary only
  // when the action is "Merge" (that's the one that keeps a surviving record).
  // Close / Evaluate / None set the group aside or defer it, so a missing
  // primary is fine for those.
  function unclassifiedComplete(g) {
    if (!g.actionChosen) return false;
    if (g.action !== 'Merge') return true;
    return g.rows.some(function (r) { return r.isPrimary; });
  }
  function shouldAutoNext(g, sheet) {
    if (sheet.key === 'unclassified') return unclassifiedComplete(g);
    return g.actionChosen;
  }
  function scheduleNext() { setTimeout(function () { navGroup(1); }, 200); }

  function setAction(a) {
    var g = currentGroup();
    g.action = a;
    g.actionChosen = true;
    updateActionButtons();
    saveState();
    // Switching between actions changes whether a primary is still required.
    if (currentSheet().key === 'unclassified') renderPrimaryName();
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
    saveState();
    renderEditGroup();
    // Unclassified auto-advances only once both an action and a primary are set.
    if (currentSheet().key === 'unclassified' && shouldAutoNext(g, currentSheet())) scheduleNext();
  }

  function deleteRecord(idx) {
    var g = currentGroup();
    if (g.rows.length <= 2 || idx < 0 || idx >= g.rows.length) return;
    var removed = g.rows.splice(idx, 1)[0];
    if (!g.deletedRows) g.deletedRows = [];
    g.deletedRows.push(removed);   // remember for undo
    saveState();
    renderEditGroup();
  }

  // Undo the most recent record removal for the current group.
  function undoDelete() {
    var g = currentGroup();
    if (!g.deletedRows || !g.deletedRows.length) return;
    g.rows.push(g.deletedRows.pop());
    saveState();
    renderEditGroup();
  }

  function updateActionButtons() {
    var action = currentGroup().action;
    Array.prototype.forEach.call(els.actionButtons.children, function (b) {
      b.classList.toggle('active', b.getAttribute('data-action') === action);
    });
  }

  // Large centered banner: the primary account name, or a prompt. Unclassified
  // groups only need a primary when the action is "Merge", so for any other
  // chosen action we don't nag for one.
  function renderPrimaryName() {
    var sheet = currentSheet();
    var g = currentGroup();
    var primaryRow = g.rows.filter(function (r) { return r.isPrimary; })[0];
    if (primaryRow) {
      els.editPrimaryName.textContent = RP.fieldValue({ src: RP.COL.accountName, kind: 'data' }, g, primaryRow, 1) || '(unnamed account)';
      els.editPrimaryName.classList.remove('no-primary');
    } else if (sheet.key === 'unclassified' && g.actionChosen && g.action !== 'Merge') {
      els.editPrimaryName.textContent = 'No primary needed';
      els.editPrimaryName.classList.remove('no-primary');
    } else {
      els.editPrimaryName.textContent = '⚠ Select a primary account';
      els.editPrimaryName.classList.add('no-primary');
    }
  }

  function renderEditGroup() {
    var sheet = currentSheet();
    var g = currentGroup();
    var idx = state.edit.group;

    els.editSheet.textContent = 'Editing: ' + sheet.title;
    els.editProgress.textContent = 'Group ' + (idx + 1) + ' of ' + sheet.groups.length;

    renderPrimaryName();

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
    // fields (no group-level Action/Notes/Remarks). Groups with 3+ records also
    // get a Remove column so records that don't belong can be dropped.
    var dataCols = RP.displayedFields(sheet).filter(function (f) { return !f.group && f.kind !== 'classification'; });
    var canDelete = g.rows.length > 2;
    var html = '<tr class="head"><td>Primary</td>' +
      dataCols.map(function (c) { return '<td>' + esc(c.label) + '</td>'; }).join('') +
      (canDelete ? '<td class="pv-del-head">Remove</td>' : '') + '</tr>';
    g.rows.forEach(function (row, ri) {
      var radio = '<td class="pv-radio"><input type="radio" name="editPrimary" data-row="' + ri + '"' +
        (row.isPrimary ? ' checked' : '') + ' aria-label="Set as primary" /></td>';
      var del = canDelete
        ? '<td class="pv-del"><button type="button" class="row-del" data-row="' + ri +
          '" aria-label="Remove this record from the group" title="Remove from group">✕</button></td>'
        : '';
      html += '<tr' + (row.isPrimary ? ' class="primary"' : '') + '>' + radio +
        dataCols.map(function (c) { return td(c, g, row, ri); }).join('') + del + '</tr>';
    });
    els.editTable.className = 'pv theme-' + sheet.theme.name;
    els.editTable.innerHTML = html;

    // Undo the last record removal (available even after the group drops to 2
    // rows, and after a reload — the removed rows are reconstructed on restore).
    var undoN = (g.deletedRows && g.deletedRows.length) || 0;
    els.editUndo.hidden = undoN === 0;
    els.editUndo.textContent = '↶ Undo remove' + (undoN > 1 ? ' (' + undoN + ')' : '');

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
    state.cfTarget = state.edit.active ? state.result.sheets[state.edit.sheet]
      : isReviewActive() ? state.reviewSheet
      : state.result.sheets[state.activeSheet];
    state.cfEditingId = null;
    els.cfTitle.textContent = 'Customize: ' + state.cfTarget.title;
    renderCfLists();
    renderCfPresets();
    els.customizeOverlay.hidden = false;
  }
  els.cfClose.addEventListener('click', function () {
    if (state.cfEditingId) commitRename();
    els.customizeOverlay.hidden = true;
  });
  els.cfRevert.addEventListener('click', function () {
    var sheet = state.cfTarget;
    sheet.fields.forEach(function (f) { f.label = f.defaultLabel; });
    state.cfEditingId = null;
    renderCfLists();
    renderCfPresets();
    rerenderActive();
    saveState();
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
    var sheet = state.cfTarget;
    var shown = sheet.fields.slice(0, sheet.displayCount);
    var hidden = sheet.fields.slice(sheet.displayCount);
    els.cfShown.innerHTML = shown.map(cfItemHtml).join('') || cfEmpty('Drag fields here to show them');
    els.cfHidden.innerHTML = hidden.map(cfItemHtml).join('') || cfEmpty('Drag fields here to hide them');
    if (state.cfEditingId) {
      var inp = els.customizeOverlay.querySelector('.cf-rename-input');
      if (inp) { inp.focus(); inp.select(); }
    }
  }

  // Preset row: one button per saved preset, plus a "Save preset" button that
  // appears only when the current columns differ from the active preset.
  function renderCfPresets() {
    var sheet = state.cfTarget;
    if (!sheet.presets) return;
    var html = sheet.presets.map(function (p, i) {
      return '<button type="button" class="cf-preset' + (i === sheet.activePreset ? ' active' : '') +
        '" data-preset="' + i + '">' + esc(p.name) + '</button>';
    }).join('');
    if (presetIsDirty(sheet)) {
      html += '<button type="button" class="cf-preset cf-preset-save" data-preset="save">＋ Save preset</button>';
    }
    els.cfPresets.innerHTML = html;
  }
  els.cfPresets.addEventListener('click', function (e) {
    var b = e.target.closest('.cf-preset'); if (!b) return;
    var which = b.getAttribute('data-preset');
    var sheet = state.cfTarget;
    if (which === 'save') saveNewPreset(sheet);
    else applyPreset(sheet, parseInt(which, 10));
    state.cfEditingId = null;
    renderCfLists();
    renderCfPresets();
    rerenderActive();
    saveState();
  });

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
    var sheet = state.cfTarget;
    var f = sheet.fields.filter(function (x) { return x.id === id; })[0];
    if (f && inp) { var v = inp.value.trim(); f.label = v || f.defaultLabel; }
    state.cfEditingId = null;
    renderCfLists();
    renderCfPresets();
    rerenderActive();
    saveState();
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
    var sheet = state.cfTarget;
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
    renderCfPresets();
    rerenderActive();
    saveState();
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
