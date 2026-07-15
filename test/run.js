/*
 * Node test suite for report.js + xlsx.js. Run with:  node test/run.js
 * Tests the sheet model (group-level Action/Remarks/Notes, detail columns,
 * primary-on-top) and that the workbook serializes to a valid ZIP.
 * Deep XLSX validation lives in a separate scratchpad script.
 */
var RP = require('../report.js');
var XW = require('../xlsx.js');
var fs = require('fs');
var path = require('path');

var failures = 0;
function assert(cond, msg) { console.log((cond ? '  ok  - ' : '  FAIL- ') + msg); if (!cond) failures++; }
function eq(a, b, msg) { assert(JSON.stringify(a) === JSON.stringify(b), msg + '  (got ' + JSON.stringify(a) + ', expected ' + JSON.stringify(b) + ')'); }

require('./gen-sample.js');
var text = fs.readFileSync(path.join(__dirname, 'sample-export.csv'), 'utf8');
var result = RP.buildReport(RP.parseCSV(text));
var likely = result.sheets[0], attention = result.sheets[1], unclassified = result.sheets[2];

console.log('\n=== Stats ===');
console.log(JSON.stringify(result.stats));

console.log('\n=== Model assertions ===');
eq(result.sheets.length, 3, 'Three sheets');
eq([likely.title, attention.title, unclassified.title], ['Likely duplicates', 'Needs Attention', 'Unclassified'], 'Sheet titles');
eq([likely.editable, attention.editable, unclassified.editable], [true, true, false], 'Editable flags (unclassified not editable)');

// Columns (Remarks added to editable sheets; unclassified has none)
eq(likely.columns.map(function (c) { return c.h; }),
  ['Action', 'Classification', 'Account Name', 'Account ID', 'Type', 'Owner', 'Reason', 'Remarks'], 'Likely columns + Remarks');
eq(attention.columns.map(function (c) { return c.h; }),
  ['Action', 'Classification', 'Account Name', 'Account ID', 'Type', 'Owner', 'Reason', 'Notes', 'Remarks'], 'Needs Attention columns');
eq(unclassified.columns.map(function (c) { return c.h; }),
  ['Action', 'Classification', 'Account Name', 'Account ID', 'Type', 'Owner', 'Notes'], 'Unclassified columns (no Remarks)');

// Detail columns exclude U(20), V(21), Z(25), and already-shown columns.
eq(likely.detailColumns.map(function (c) { return c.h; }), ['Region', 'MID', 'Created', 'Relationships', 'Country Conflict'], 'Likely detail columns');
eq(unclassified.detailColumns.map(function (c) { return c.h; }), ['Region', 'MID', 'Created', 'Reason', 'Relationships', 'Country Conflict'], 'Unclassified detail includes Reason (index order)');
var allDetailIdx = likely.detailColumns.concat(attention.detailColumns, unclassified.detailColumns).map(function (c) { return c.index; });
assert(allDetailIdx.indexOf(20) === -1 && allDetailIdx.indexOf(21) === -1 && allDetailIdx.indexOf(25) === -1, 'Detail never includes U/V/Z');
assert(allDetailIdx.indexOf(23) === -1, 'Detail never includes Classification (already shown)');

// Group-level fields
var acme = likely.groups[0];
eq(acme.action, 'None', 'Group action defaults to None');
eq(acme.remarks, '', 'Group remarks default empty');
assert(acme.rows[0].isPrimary === true, 'Primary is top row');
assert(acme.rows[0].raw && acme.rows[0].raw.length > 0, 'Record carries raw row for detailed view');

// Notes at group level
eq(attention.groups[1].notes, 'has country conflict, VAT number mismatch', 'Delta group note (conflict + complexity extra)');
eq(attention.groups[2].rows[0].id, 'ACC-040', 'Epsilon primary reordered to top');
eq(attention.groups[2].notes, 'has relationship, has country conflict, needs finance sign-off', 'Epsilon combined note');
eq(unclassified.groups[0].notes, 'has relationship', 'Zeta (unclassified) note = relationship only');

// Every classified group has primary on top
[likely, attention].forEach(function (sh) {
  sh.groups.forEach(function (g, gi) { assert(g.rows[0].isPrimary === true, sh.title + ' group ' + (gi + 1) + ' primary on top'); });
});

eq(result.warnings.length, 0, 'No warnings for clean dataset');

// --- Workbook serialization (with an edited action + remark) ---------------
console.log('\n=== Workbook ===');
likely.groups[0].action = 'Merge';
likely.groups[0].remarks = 'confirmed dup of ACC-001';
attention.groups[2].action = 'Evaluate';
var bytes = XW.buildWorkbook(result.sheets);
assert(bytes instanceof Uint8Array && bytes.length > 0, 'buildWorkbook returns non-empty Uint8Array');
assert(bytes[0] === 0x50 && bytes[1] === 0x4B, 'ZIP signature "PK"');
var hasEocd = false;
for (var i = bytes.length - 22; i >= 0 && i > bytes.length - 200; i--) {
  if (bytes[i] === 0x50 && bytes[i + 1] === 0x4B && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) { hasEocd = true; break; }
}
assert(hasEocd, 'Has End-Of-Central-Directory record');
fs.writeFileSync(path.join(__dirname, 'sample-report.xlsx'), bytes);
console.log('  wrote test/sample-report.xlsx (' + bytes.length + ' bytes)');

console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
