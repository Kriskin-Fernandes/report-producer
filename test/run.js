/*
 * Node test suite for report.js + xlsx.js. Run with:  node test/run.js
 * Tests the field-universe model (displayed vs hidden, Alias, editable
 * Unclassified, field-value resolution), customization affecting export, and
 * that the workbook serializes to a valid ZIP. Deep XLSX validation lives in a
 * separate scratchpad script.
 */
var RP = require('../report.js');
var XW = require('../xlsx.js');
var fs = require('fs');
var path = require('path');

var failures = 0;
function assert(cond, msg) { console.log((cond ? '  ok  - ' : '  FAIL- ') + msg); if (!cond) failures++; }
function eq(a, b, msg) { assert(JSON.stringify(a) === JSON.stringify(b), msg + '  (got ' + JSON.stringify(a) + ', expected ' + JSON.stringify(b) + ')'); }
function labels(fields) { return fields.map(function (f) { return f.label; }); }

require('./gen-sample.js');
var text = fs.readFileSync(path.join(__dirname, 'sample-export.csv'), 'utf8');
var result = RP.buildReport(RP.parseCSV(text));
var likely = result.sheets[0], attention = result.sheets[1], unclassified = result.sheets[2];

console.log('\n=== Stats ===');
console.log(JSON.stringify(result.stats));

console.log('\n=== Model assertions ===');
eq([likely.title, attention.title, unclassified.title],
  ['Likely duplicates', 'Problematic duplicates', 'Unclassified'], 'Sheet titles (renamed)');
eq([likely.editable, attention.editable, unclassified.editable], [true, true, true], 'All sheets editable');

// Displayed fields, with Alias right of Account Name.
eq(labels(RP.displayedFields(likely)),
  ['Action', 'Classification', 'Account Name', 'Alias account', 'Account ID', 'Type', 'Owner', 'Reason', 'Remarks'],
  'Likely displayed fields (Alias after Account Name)');
eq(labels(RP.displayedFields(attention)),
  ['Action', 'Classification', 'Account Name', 'Alias account', 'Account ID', 'Type', 'Owner', 'Reason', 'Notes', 'Remarks'],
  'Problematic displayed fields');
eq(labels(RP.displayedFields(unclassified)),
  ['Action', 'Classification', 'Account Name', 'Alias account', 'Account ID', 'Type', 'Owner', 'Notes', 'Remarks'],
  'Unclassified displayed fields (has Remarks now)');

// Hidden (detail) fields.
eq(labels(likely.fields.slice(likely.displayCount)),
  ['Region', 'MID', 'Created', 'Relationships', 'Country Conflict'], 'Likely hidden fields');
assert(labels(unclassified.fields.slice(unclassified.displayCount)).indexOf('Reason') >= 0,
  'Unclassified hidden fields include Reason (not displayed there)');

// Never-show columns U/V/Z never appear as a field src anywhere.
var allSrcs = [];
result.sheets.forEach(function (s) { s.fields.forEach(function (f) { if (f.src != null) allSrcs.push(f.src); }); });
assert(allSrcs.indexOf(20) === -1 && allSrcs.indexOf(21) === -1 && allSrcs.indexOf(25) === -1, 'No field sources in U/V/Z');

// Alias field maps to column B.
var aliasField = RP.displayedFields(likely).filter(function (f) { return f.id === 'alias'; })[0];
eq(aliasField.src, RP.COL.alias, 'Alias field source = column B');

// ID field is a monospace Salesforce link.
var idField = RP.displayedFields(likely).filter(function (f) { return f.id === 'id'; })[0];
assert(idField.link === true && idField.mono === true, 'Account ID field is link + monospace');

console.log('\n=== Field value resolution ===');
var acme = likely.groups[0];
eq(acme.action, 'None', 'Group action default None');
eq(acme.actionChosen, false, 'actionChosen default false');
assert(acme.rows[0].isPrimary === true, 'Primary is top row');
var nameF = RP.displayedFields(likely).filter(function (f) { return f.id === 'name'; })[0];
var idF = RP.displayedFields(likely).filter(function (f) { return f.id === 'id'; })[0];
var aliasF = RP.displayedFields(likely).filter(function (f) { return f.id === 'alias'; })[0];
var classF = RP.displayedFields(likely).filter(function (f) { return f.id === 'classification'; })[0];
var actionF = RP.displayedFields(likely).filter(function (f) { return f.id === 'action'; })[0];
eq(RP.fieldValue(nameF, acme, acme.rows[0], 0), 'Acme, Inc.', 'name field value');
eq(RP.fieldValue(idF, acme, acme.rows[0], 0), 'ACC-001', 'id field value');
eq(RP.fieldValue(aliasF, acme, acme.rows[0], 0), 'ACME Holdings', 'alias field value (col B)');
eq(RP.fieldValue(classF, acme, acme.rows[0], 0), 'Primary', 'classification value from row');
eq(RP.fieldValue(actionF, acme, acme.rows[0], 0), 'None', 'action shows on top row');
eq(RP.fieldValue(actionF, acme, acme.rows[1], 1), '', 'action blank on duplicate row');

// Notes at group level
eq(attention.groups[1].notes, 'has country conflict, VAT number mismatch', 'Delta notes');
eq(attention.groups[2].rows[0].isPrimary && RP.fieldValue(idF, attention.groups[2], attention.groups[2].rows[0], 0), 'ACC-040', 'Epsilon primary on top');

console.log('\n=== Customization affects export ===');
// Move "Owner" below the cutoff on Likely -> it should drop from displayed.
var likelyFields = likely.fields;
var ownerIdx = likelyFields.findIndex(function (f) { return f.id === 'owner'; });
var owner = likelyFields.splice(ownerIdx, 1)[0];
likelyFields.push(owner);           // move to the very end (hidden region)
likely.displayCount -= 1;           // shrink the cutoff by one
assert(labels(RP.displayedFields(likely)).indexOf('Owner') === -1, 'Owner removed from displayed after customize');

console.log('\n=== Workbook ===');
attention.groups[0].action = 'Merge';
attention.groups[0].actionChosen = true;
attention.groups[0].remarks = 'reviewed';
var bytes = XW.buildWorkbook(result.sheets);
assert(bytes instanceof Uint8Array && bytes.length > 0, 'buildWorkbook returns non-empty Uint8Array');
assert(bytes[0] === 0x50 && bytes[1] === 0x4B, 'ZIP signature "PK"');
var hasEocd = false;
for (var i = bytes.length - 22; i >= 0 && i > bytes.length - 200; i--) {
  if (bytes[i] === 0x50 && bytes[i + 1] === 0x4B && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) { hasEocd = true; break; }
}
assert(hasEocd, 'Has End-Of-Central-Directory record');

// Rebuild from a FRESH parse so the saved sample .xlsx reflects the default
// (un-customized) layout for the deep validator.
var fresh = RP.buildReport(RP.parseCSV(text));
fresh.sheets[0].groups[0].action = 'Merge';
fresh.sheets[0].groups[0].remarks = 'confirmed dup of ACC-001';
fs.writeFileSync(path.join(__dirname, 'sample-report.xlsx'), XW.buildWorkbook(fresh.sheets));
console.log('  wrote test/sample-report.xlsx');

console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
