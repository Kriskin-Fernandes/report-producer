/*
 * Node test suite for report.js + xlsx.js. Run with:  node test/run.js
 * Tests the sheet model and that the workbook serializes to a valid ZIP.
 * Deep XLSX validation (SheetJS / ExcelJS read-back) lives in a separate
 * script run from the scratchpad where those libs are installed.
 */
var RP = require('../report.js');
var XW = require('../xlsx.js');
var fs = require('fs');
var path = require('path');

var failures = 0;
function assert(cond, msg) {
  console.log((cond ? '  ok  - ' : '  FAIL- ') + msg);
  if (!cond) failures++;
}
function eq(a, b, msg) {
  assert(JSON.stringify(a) === JSON.stringify(b), msg + '  (got ' + JSON.stringify(a) + ', expected ' + JSON.stringify(b) + ')');
}

// Regenerate the sample and parse it.
require('./gen-sample.js');
var text = fs.readFileSync(path.join(__dirname, 'sample-export.csv'), 'utf8');
var result = RP.buildReport(RP.parseCSV(text));

console.log('\n=== Stats ===');
console.log(JSON.stringify(result.stats, null, 2));
console.log('=== Warnings ===');
console.log(JSON.stringify(result.warnings));

var sheets = result.sheets;
var likely = sheets[0], attention = sheets[1], unclassified = sheets[2];

console.log('\n=== Model assertions ===');
eq(sheets.length, 3, 'Three sheets');
eq([likely.title, attention.title, unclassified.title],
  ['Likely duplicates', 'Needs Attention', 'Unclassified'], 'Sheet titles');
eq([likely.theme.name, attention.theme.name, unclassified.theme.name],
  ['blue', 'peach', 'gray'], 'Sheet themes');

// Section membership
eq(result.stats.section1Groups, 2, 'Section 1 has 2 groups (Acme, Beta)');
eq(result.stats.section2Groups, 3, 'Section 2 has 3 groups (Gamma, Delta, Epsilon)');
eq(result.stats.unclassifiedGroups, 2, 'Unclassified has 2 groups (Zeta, Eta)');
eq(result.stats.section1Records, 5, 'Section 1 record count');
eq(result.stats.unclassifiedRecords, 4, 'Unclassified record count');

// Columns
eq(likely.columns.map(function (c) { return c.h; }),
  ['Action', 'Classification', 'Account Name', 'Account ID', 'Type', 'Owner', 'Reason'],
  'Likely columns');
eq(attention.columns.map(function (c) { return c.h; }),
  ['Action', 'Classification', 'Account Name', 'Account ID', 'Type', 'Owner', 'Reason', 'Notes'],
  'Needs Attention columns (with Notes)');
eq(unclassified.columns.map(function (c) { return c.h; }),
  ['Action', 'Classification', 'Account Name', 'Account ID', 'Type', 'Owner', 'Notes'],
  'Unclassified columns (Notes, no Reason)');

// Acme group in section 1: 3 rows, first is primary (bold), Action default None
var acme = likely.groups[0];
eq(acme.rows.length, 3, 'Acme group has 3 rows');
assert(acme.rows[0].isPrimary === true, 'Acme first row is primary');
assert(acme.rows[1].isPrimary === false, 'Acme second row is duplicate');
eq(acme.rows[0].action, 'None', 'Action defaults to None');
eq(acme.rows[0].classification, 'Primary', 'Classification carried over');
eq(acme.rows[0].id, 'ACC-001', 'Account ID present');

// Beta group: Low with "- clean match" complexity note, no rel/conflict -> section 1
var beta = likely.groups[1];
eq(beta.rows[0].name, 'Beta Co', 'Beta primary in section 1 (Low despite complexity note)');

// Gamma (section 2): relationship note on primary row
var gamma = attention.groups[0];
var gammaPrimary = gamma.rows.filter(function (r) { return r.isPrimary; })[0];
eq(gammaPrimary.notes, 'has relationship', 'Gamma notes = has relationship');
eq(gamma.rows.filter(function (r) { return !r.isPrimary; })[0].notes, '', 'Non-primary rows have no notes');

// Delta (section 2): conflict + complexity extra
var delta = attention.groups[1];
var deltaPrimary = delta.rows.filter(function (r) { return r.isPrimary; })[0];
eq(deltaPrimary.notes, 'has country conflict, VAT number mismatch', 'Delta notes = conflict + complexity extra');

// Epsilon (section 2): primary was second in the input but is reordered to the
// top; rel + conflict + complexity extra on the (now top) primary row.
var epsilon = attention.groups[2];
assert(epsilon.rows[0].isPrimary === true && epsilon.rows[1].isPrimary === false, 'Epsilon primary reordered to top row');
eq(epsilon.rows[0].id, 'ACC-040', 'Epsilon top row is the primary account (ACC-040)');
eq(epsilon.rows[0].notes, 'has relationship, has country conflict, needs finance sign-off', 'Epsilon combined notes on top row');
eq(epsilon.rows[1].notes, '', 'Epsilon non-primary note empty');

// Every classified group has its primary as the top row.
[likely, attention].forEach(function (sh) {
  sh.groups.forEach(function (g, gi) {
    assert(g.rows[0].isPrimary === true, sh.title + ' group ' + (gi + 1) + ' has primary on top row');
  });
});

// Unclassified: Zeta group of 3, no primaries, relationship note on first row
var zeta = unclassified.groups[0];
eq(zeta.rows.length, 3, 'Zeta has 3 rows');
assert(zeta.rows.every(function (r) { return !r.isPrimary; }), 'No primaries in unclassified');
eq(zeta.rows[0].notes, 'has relationship', 'Zeta first-row note (rel only, no complexity)');
eq(zeta.rows[0].classification, 'Unclassified', 'Unclassified classification carried');

eq(result.warnings.length, 0, 'No warnings for clean dataset');

// --- Workbook serialization ------------------------------------------------
console.log('\n=== Workbook bytes ===');
var bytes = XW.buildWorkbook(sheets);
assert(bytes instanceof Uint8Array && bytes.length > 0, 'buildWorkbook returns non-empty Uint8Array');
assert(bytes[0] === 0x50 && bytes[1] === 0x4B, 'Starts with ZIP signature "PK"');
// EOCD signature present near the end
var hasEocd = false;
for (var i = bytes.length - 22; i >= 0 && i > bytes.length - 200; i--) {
  if (bytes[i] === 0x50 && bytes[i + 1] === 0x4B && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) { hasEocd = true; break; }
}
assert(hasEocd, 'Contains End-Of-Central-Directory record');

var outPath = path.join(__dirname, 'sample-report.xlsx');
fs.writeFileSync(outPath, bytes);
console.log('  wrote ' + outPath + ' (' + bytes.length + ' bytes)');

console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
