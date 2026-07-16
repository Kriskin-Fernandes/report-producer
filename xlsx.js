/*
 * xlsx.js — a tiny, dependency-free XLSX (OOXML SpreadsheetML) writer.
 *
 * Builds a multi-sheet .xlsx workbook entirely in memory and returns the raw
 * bytes (Uint8Array). No network, no third-party library. Supports exactly the
 * features this report needs: merged+centered cells, per-cell fonts (size /
 * bold / italic / monospace), solid fills, borders (full grid, group boxes,
 * column rules), a list data-validation dropdown, and frozen header rows.
 *
 * Public API:
 *   XlsxWriter.buildWorkbook(sheets) -> Uint8Array
 * where `sheets` is the model produced by report.js buildReport().sheets.
 */
(function (global) {
  'use strict';

  var MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  var CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

  var enc = new TextEncoder();

  // ---- XML helpers ----------------------------------------------------------
  function xmlEscape(s) {
    return String(s)
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '') // strip invalid XML control chars
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function colLetter(n) { // 1-based -> A, B, ... Z, AA...
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }

  // ---- Style manager (dedups fonts / fills / borders / cellXfs) -------------
  function StyleManager() {
    // Fonts: index 0 must be the default.
    this.fonts = ['<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'];
    // Fills: indexes 0 and 1 are reserved by the spec (none, gray125).
    this.fills = [
      '<fill><patternFill patternType="none"/></fill>',
      '<fill><patternFill patternType="gray125"/></fill>'
    ];
    // Borders: index 0 = no borders.
    this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    // cellXfs: index 0 = default (plain).
    this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    this._fontKey = {}; this._fillKey = {}; this._borderKey = {}; this._xfKey = {};
  }

  StyleManager.prototype.font = function (spec) {
    // spec: {size, bold, italic, mono, underline, color}
    var key = JSON.stringify(spec);
    if (this._fontKey[key] != null) return this._fontKey[key];
    var name = spec.mono ? 'Consolas' : 'Calibri';
    var family = spec.mono ? 3 : 2;
    var xml = '<font>' +
      (spec.bold ? '<b/>' : '') +
      (spec.italic ? '<i/>' : '') +
      (spec.underline ? '<u/>' : '') +
      '<sz val="' + (spec.size || 11) + '"/>' +
      (spec.color ? '<color rgb="' + spec.color + '"/>' : '') +
      '<name val="' + name + '"/>' +
      '<family val="' + family + '"/>' +
      '</font>';
    var id = this.fonts.length;
    this.fonts.push(xml);
    this._fontKey[key] = id;
    return id;
  };

  StyleManager.prototype.solidFill = function (rgb) {
    if (!rgb) return 0;
    var key = 'S' + rgb;
    if (this._fillKey[key] != null) return this._fillKey[key];
    var xml = '<fill><patternFill patternType="solid"><fgColor rgb="FF' + rgb +
      '"/><bgColor indexed="64"/></patternFill></fill>';
    var id = this.fills.length;
    this.fills.push(xml);
    this._fillKey[key] = id;
    return id;
  };

  StyleManager.prototype.border = function (spec) {
    // spec: {l, r, t, b} booleans; thin black.
    var key = (spec.l ? 1 : 0) + '' + (spec.r ? 1 : 0) + (spec.t ? 1 : 0) + (spec.b ? 1 : 0);
    if (this._borderKey[key] != null) return this._borderKey[key];
    function side(name, on) {
      return on ? '<' + name + ' style="thin"><color rgb="FF000000"/></' + name + '>' : '<' + name + '/>';
    }
    var xml = '<border>' + side('left', spec.l) + side('right', spec.r) +
      side('top', spec.t) + side('bottom', spec.b) + '<diagonal/></border>';
    var id = this.borders.length;
    this.borders.push(xml);
    this._borderKey[key] = id;
    return id;
  };

  StyleManager.prototype.xf = function (o) {
    // o: {fontId, fillId, borderId, halign, valign, wrap}
    var fontId = o.fontId || 0, fillId = o.fillId || 0, borderId = o.borderId || 0;
    var halign = o.halign || '', valign = o.valign || '', wrap = !!o.wrap;
    var key = fontId + ':' + fillId + ':' + borderId + ':' + halign + ':' + valign + ':' + (wrap ? 1 : 0);
    if (this._xfKey[key] != null) return this._xfKey[key];
    var hasAlign = halign || valign || wrap;
    var xml = '<xf numFmtId="0" fontId="' + fontId + '" fillId="' + fillId +
      '" borderId="' + borderId + '" xfId="0"' +
      (fontId ? ' applyFont="1"' : '') +
      (fillId ? ' applyFill="1"' : '') +
      (borderId ? ' applyBorder="1"' : '') +
      (hasAlign ? ' applyAlignment="1"' : '') + '>';
    if (hasAlign) {
      xml += '<alignment' +
        (halign ? ' horizontal="' + halign + '"' : '') +
        (valign ? ' vertical="' + valign + '"' : '') +
        (wrap ? ' wrapText="1"' : '') + '/>';
    }
    xml += '</xf>';
    var id = this.xfs.length;
    this.xfs.push(xml);
    this._xfKey[key] = id;
    return id;
  };

  StyleManager.prototype.toXml = function () {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="' + MAIN_NS + '">' +
      '<fonts count="' + this.fonts.length + '">' + this.fonts.join('') + '</fonts>' +
      '<fills count="' + this.fills.length + '">' + this.fills.join('') + '</fills>' +
      '<borders count="' + this.borders.length + '">' + this.borders.join('') + '</borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + this.xfs.length + '">' + this.xfs.join('') + '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>';
  };

  var SALESFORCE_BASE = 'https://checkout.my.salesforce.com/';

  // Resolve a field's value for a row (rowIndex within its group). Group-level
  // fields (action / notes / remarks) render on the top row only.
  function fieldValue(f, group, row, ri) {
    if (f.group) return ri === 0 ? (group[f.id] || '') : '';
    if (f.kind === 'classification') return row.classification || '';
    var v = row.raw ? row.raw[f.src] : '';
    return v == null ? '' : String(v).trim();
  }

  // ---- Worksheet XML --------------------------------------------------------
  function sheetXml(sheet, sm) {
    var cols = sheet.fields.slice(0, sheet.displayCount); // displayed fields only
    var nCols = cols.length;
    var lastCol = colLetter(nCols);
    var theme = sheet.theme;

    var headerFillId = sm.solidFill(theme.headerFill);
    var allBorder = sm.border({ l: 1, r: 1, t: 1, b: 1 });

    // Header / description / column-heading styles.
    var xfTitle = sm.xf({ fontId: sm.font({ size: 20, bold: true }), fillId: headerFillId, borderId: allBorder, halign: 'center', valign: 'center' });
    var xfTitleFill = sm.xf({ fillId: headerFillId, borderId: allBorder });
    var xfDesc = sm.xf({ fontId: sm.font({ size: 15, italic: true }), fillId: headerFillId, borderId: allBorder, halign: 'center', valign: 'center' });
    var xfDescFill = sm.xf({ fillId: headerFillId, borderId: allBorder });
    var xfHeadCell = sm.xf({ fontId: sm.font({ size: 11, bold: true }), fillId: headerFillId, borderId: allBorder, halign: 'center', valign: 'center' });

    // Data-cell style resolver (dedups automatically via the manager).
    // Primary rows are bold but carry no background fill. Linked (ID) cells are
    // blue + underlined.
    function dataXf(field, isPrimary, top, bottom) {
      var fontId = sm.font({
        size: 11,
        bold: !!isPrimary,
        mono: !!field.mono,
        underline: !!field.link,
        color: field.link ? 'FF0563C1' : null
      });
      var borderId = sm.border({ l: 1, r: 1, t: !!top, b: !!bottom });
      return sm.xf({
        fontId: fontId, fillId: 0, borderId: borderId,
        halign: field.align || 'left', valign: 'center', wrap: !!field.wrap
      });
    }

    var rowsXml = [];
    var rowNum = 0;
    function pushRow(r, cells, ht) {
      rowsXml.push('<row r="' + r + '"' + (ht ? ' ht="' + ht + '" customHeight="1"' : '') + '>' + cells + '</row>');
    }
    
    function styledCell(ref, s, value, isLink) {
      if (value == null || value === '') return '<c r="' + ref + '" s="' + s + '"/>';

      // Linked (Account ID) cells become a Salesforce HYPERLINK formula.
      if (isLink) {
        var safeVal = xmlEscape(value);
        var excelSafeStr = String(value).replace(/"/g, '""'); // escape inner quotes for the formula
        var formula = 'HYPERLINK("' + SALESFORCE_BASE + excelSafeStr + '", "' + excelSafeStr + '")';
        // t="str" => <v> holds the formula's cached result.
        return '<c r="' + ref + '" s="' + s + '" t="str"><f>' + xmlEscape(formula) + '</f><v>' + safeVal + '</v></c>';
      }

      return '<c r="' + ref + '" s="' + s + '" t="inlineStr"><is><t xml:space="preserve">' +
        xmlEscape(value) + '</t></is></c>';
    }

    // Row 1: merged section title.
    rowNum = 1;
    var titleCells = '';
    for (var ci = 1; ci <= nCols; ci++) {
      titleCells += styledCell(colLetter(ci) + '1', ci === 1 ? xfTitle : xfTitleFill, ci === 1 ? sheet.title : '');
    }
    pushRow(1, titleCells, 30);

    // Row 2: merged description.
    var descCells = '';
    for (var cd = 1; cd <= nCols; cd++) {
      descCells += styledCell(colLetter(cd) + '2', cd === 1 ? xfDesc : xfDescFill, cd === 1 ? sheet.description : '');
    }
    pushRow(2, descCells, 22);

    // Row 3: column headings.
    var headCells = '';
    for (var ch = 0; ch < nCols; ch++) {
      headCells += styledCell(colLetter(ch + 1) + '3', xfHeadCell, cols[ch].label);
    }
    pushRow(3, headCells, 18);

    // Which displayed column is the Action column (dropdown target)?
    var actionIdx = -1;
    for (var ai = 0; ai < nCols; ai++) { if (cols[ai].id === 'action') { actionIdx = ai; break; } }

    // Data groups (blank row between groups) + collect dropdown cells.
    rowNum = 3;
    var validationRanges = [];
    sheet.groups.forEach(function (group, gi) {
      if (gi > 0) rowNum++; // blank separator row (simply skip a row number)
      var gRows = group.rows;
      var startRow = rowNum + 1;
      gRows.forEach(function (rec, ri) {
        rowNum++;
        var top = ri === 0;
        var bottom = ri === gRows.length - 1;
        var cellsXml = '';
        for (var c = 0; c < nCols; c++) {
          var field = cols[c];
          var val = fieldValue(field, group, rec, ri);
          var s = dataXf(field, rec.isPrimary, top, bottom);
          cellsXml += styledCell(colLetter(c + 1) + rowNum, s, val, field.link);
        }
        pushRow(rowNum, cellsXml);
      });
      // Action dropdown lives on the group's top row only.
      if (actionIdx >= 0) validationRanges.push(colLetter(actionIdx + 1) + startRow);
    });

    var lastRow = rowNum < 3 ? 3 : rowNum;

    // <cols> widths. Fit-fields (Notes / Remarks) do not wrap, so widen them to
    // fit their longest (group-level) value and keep every value on one line.
    function fitWidth(field) {
      var maxLen = field.label.length;
      sheet.groups.forEach(function (g) {
        var s = g[field.id] == null ? '' : String(g[field.id]);
        if (s.length > maxLen) maxLen = s.length;
      });
      return Math.min(120, Math.max(field.width || 16, maxLen + 3));
    }
    var colsXml = '<cols>';
    for (var cw = 0; cw < nCols; cw++) {
      var f = cols[cw];
      var w = f.fit ? fitWidth(f) : (f.width || 16);
      colsXml += '<col min="' + (cw + 1) + '" max="' + (cw + 1) + '" width="' + w + '" customWidth="1"/>';
    }
    colsXml += '</cols>';

    var mergeXml = '<mergeCells count="2"><mergeCell ref="A1:' + lastCol + '1"/>' +
      '<mergeCell ref="A2:' + lastCol + '2"/></mergeCells>';

    var validationXml = '';
    if (validationRanges.length) {
      var opts = '&quot;' + sheet.actionOptions.join(',') + '&quot;';
      validationXml = '<dataValidations count="1">' +
        '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="' +
        validationRanges.join(' ') + '"><formula1>' + opts + '</formula1></dataValidation>' +
        '</dataValidations>';
    }

    // Freeze the three header rows.
    var viewsXml = '<sheetViews><sheetView workbookViewId="0">' +
      '<pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>' +
      '<selection pane="bottomLeft" activeCell="A4" sqref="A4"/></sheetView></sheetViews>';

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="' + MAIN_NS + '" xmlns:r="' + REL_NS + '">' +
      '<dimension ref="A1:' + lastCol + lastRow + '"/>' +
      viewsXml +
      '<sheetFormatPr defaultRowHeight="15"/>' +
      colsXml +
      '<sheetData>' + rowsXml.join('') + '</sheetData>' +
      mergeXml +
      validationXml +
      '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
      '</worksheet>';
  }

  // ---- Package assembly -----------------------------------------------------
  function buildWorkbook(sheets) {
    var sm = new StyleManager();
    var sheetXmls = sheets.map(function (s) { return sheetXml(s, sm); });
    var stylesXmlStr = sm.toXml();

    var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="' + CT_NS + '">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheets.map(function (s, i) {
        return '<Override PartName="/xl/worksheets/sheet' + (i + 1) +
          '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      }).join('') +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>';

    var rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="' + PKG_REL_NS + '">' +
      '<Relationship Id="rId1" Type="' + REL_NS + '/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>';

    var sheetTags = sheets.map(function (s, i) {
      return '<sheet name="' + xmlEscape(sheetName(s.title)) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>';
    }).join('');
    var workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="' + MAIN_NS + '" xmlns:r="' + REL_NS + '">' +
      '<sheets>' + sheetTags + '</sheets></workbook>';

    var wbRelParts = sheets.map(function (s, i) {
      return '<Relationship Id="rId' + (i + 1) + '" Type="' + REL_NS +
        '/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>';
    }).join('');
    var stylesRelId = 'rId' + (sheets.length + 1);
    var workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="' + PKG_REL_NS + '">' + wbRelParts +
      '<Relationship Id="' + stylesRelId + '" Type="' + REL_NS + '/styles" Target="styles.xml"/>' +
      '</Relationships>';

    var files = [
      { name: '[Content_Types].xml', data: contentTypes },
      { name: '_rels/.rels', data: rootRels },
      { name: 'xl/workbook.xml', data: workbook },
      { name: 'xl/_rels/workbook.xml.rels', data: workbookRels },
      { name: 'xl/styles.xml', data: stylesXmlStr }
    ];
    sheetXmls.forEach(function (x, i) {
      files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: x });
    });

    return zipStore(files);
  }

  function sheetName(title) {
    // Excel tab names: <=31 chars, none of : \ / ? * [ ]
    return String(title).replace(/[:\\\/\?\*\[\]]/g, ' ').slice(0, 31);
  }

  // ---- ZIP (store / no compression) -----------------------------------------
  var CRC_TABLE = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function zipStore(files) {
    // Fixed DOS timestamp (1980-01-01 00:00) keeps output deterministic and
    // avoids any reliance on the clock.
    var DOS_TIME = 0, DOS_DATE = 0x21;
    var chunks = [];
    var central = [];
    var offset = 0;

    function u16(a, v) { a.push(v & 0xFF, (v >>> 8) & 0xFF); }
    function u32(a, v) { a.push(v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF); }

    files.forEach(function (f) {
      var nameBytes = enc.encode(f.name);
      var dataBytes = f.data instanceof Uint8Array ? f.data : enc.encode(f.data);
      var crc = crc32(dataBytes);
      var size = dataBytes.length;

      var local = [];
      u32(local, 0x04034b50);
      u16(local, 20);          // version needed
      u16(local, 0);           // flags
      u16(local, 0);           // method: store
      u16(local, DOS_TIME);
      u16(local, DOS_DATE);
      u32(local, crc);
      u32(local, size);        // compressed size
      u32(local, size);        // uncompressed size
      u16(local, nameBytes.length);
      u16(local, 0);           // extra len
      var localHeader = Uint8Array.from(local);

      chunks.push(localHeader, nameBytes, dataBytes);

      var cen = [];
      u32(cen, 0x02014b50);
      u16(cen, 20);            // version made by
      u16(cen, 20);            // version needed
      u16(cen, 0);             // flags
      u16(cen, 0);             // method
      u16(cen, DOS_TIME);
      u16(cen, DOS_DATE);
      u32(cen, crc);
      u32(cen, size);
      u32(cen, size);
      u16(cen, nameBytes.length);
      u16(cen, 0);             // extra
      u16(cen, 0);             // comment
      u16(cen, 0);             // disk number start
      u16(cen, 0);             // internal attrs
      u32(cen, 0);             // external attrs
      u32(cen, offset);        // local header offset
      central.push(Uint8Array.from(cen), nameBytes);

      offset += localHeader.length + nameBytes.length + dataBytes.length;
    });

    var centralStart = offset;
    var centralSize = 0;
    central.forEach(function (c) { centralSize += c.length; });

    var eocd = [];
    u32(eocd, 0x06054b50);
    u16(eocd, 0);              // disk number
    u16(eocd, 0);              // disk with central dir
    u16(eocd, files.length);   // entries on this disk
    u16(eocd, files.length);   // total entries
    u32(eocd, centralSize);
    u32(eocd, centralStart);
    u16(eocd, 0);              // comment length

    var all = chunks.concat(central, [Uint8Array.from(eocd)]);
    var total = all.reduce(function (n, c) { return n + c.length; }, 0);
    var out = new Uint8Array(total);
    var pos = 0;
    all.forEach(function (c) { out.set(c, pos); pos += c.length; });
    return out;
  }

  var api = { buildWorkbook: buildWorkbook, colLetter: colLetter };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  global.XlsxWriter = api;
})(typeof self !== 'undefined' ? self : (typeof globalThis !== 'undefined' ? globalThis : this));
