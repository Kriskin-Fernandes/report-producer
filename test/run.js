/*
 * Node test suite for report.js. Run with:  node test/run.js
 * Exercises every section, the Valid-website logic, Notes, grouping,
 * and the CSV parser round-trip. No browser required.
 */
var RP = require('../report.js');

var failures = 0;
function assert(cond, msg) {
  if (cond) {
    console.log('  ok  - ' + msg);
  } else {
    failures++;
    console.error('  FAIL- ' + msg);
  }
}
function eq(a, b, msg) {
  var pa = JSON.stringify(a);
  var pb = JSON.stringify(b);
  assert(pa === pb, msg + '  (got ' + pa + ', expected ' + pb + ')');
}

// Build a 29-cell record from an overrides object keyed by column field name.
function mk(overrides) {
  var row = [];
  for (var i = 0; i < 29; i++) row.push('');
  Object.keys(overrides).forEach(function (k) {
    row[RP.COL[k]] = overrides[k];
  });
  return row;
}

// --- Header row (names just for readability / mapping panel) ---------------
var header = [];
for (var i = 0; i < 29; i++) header.push('col' + i);
header[RP.COL.primaryName] = 'Account Name';
header[RP.COL.accountId] = 'Account ID';
header[RP.COL.group] = 'Stripped Domain';
header[RP.COL.classification] = 'Classification';

var rows = [header];

// Group G1: Section 1 — classified, Low, no rel/conf, primary W=Y -> "Yes"
rows.push(mk({ primaryName: 'Acme', accountId: 'A1', type: 'Standard', owner: 'Alice',
  group: 'acme.com', validWebsite: 'Y', classification: 'Primary', reason: 'Same domain',
  complexity: 'Low', dupeProbability: '0.95' }));
rows.push(mk({ primaryName: 'Acme Dup', accountId: 'A2', group: 'acme.com',
  validWebsite: 'N', classification: 'Duplicate', complexity: 'Low' }));
rows.push(mk({ primaryName: 'Acme Dup2', accountId: 'A3', group: 'acme.com',
  validWebsite: 'N', classification: 'Duplicate', complexity: 'Low' }));

// Group G2: Section 1 — primary N, a dupe Y -> "Yes-from dupe"
rows.push(mk({ primaryName: 'Beta', accountId: 'B1', group: 'beta.com',
  validWebsite: 'N', classification: 'Primary', complexity: 'Low', dupeProbability: '0.8' }));
rows.push(mk({ primaryName: 'Beta Dup', accountId: 'B2', group: 'beta.com',
  validWebsite: 'Y', classification: 'Duplicate', complexity: 'Low' }));

// Group G3: Section 1 — all N -> "No"
rows.push(mk({ primaryName: 'Gamma', accountId: 'C1', group: 'gamma.com',
  validWebsite: 'N', classification: 'Primary', complexity: 'Low' }));
rows.push(mk({ primaryName: 'Gamma Dup', accountId: 'C2', group: 'gamma.com',
  validWebsite: 'N', classification: 'Duplicate', complexity: 'Low' }));

// Group G4: Section 2 — has relationship (on a duplicate row)
rows.push(mk({ primaryName: 'Delta', accountId: 'D1', group: 'delta.com',
  validWebsite: 'Y', classification: 'Primary', complexity: 'Low' }));
rows.push(mk({ primaryName: 'Delta Dup', accountId: 'D2', group: 'delta.com',
  classification: 'Duplicate', complexity: 'Low', relationships: 'parent:X' }));

// Group G5: Section 2 — has country conflict
rows.push(mk({ primaryName: 'Epsilon', accountId: 'E1', group: 'epsilon.com',
  validWebsite: 'Y', classification: 'Primary', complexity: 'Low', conflicts: 'GB vs US' }));
rows.push(mk({ primaryName: 'Epsilon Dup', accountId: 'E2', group: 'epsilon.com',
  classification: 'Duplicate', complexity: 'Low' }));

// Group G6: Section 2 — both relationship & conflict
rows.push(mk({ primaryName: 'Zeta', accountId: 'F1', group: 'zeta.com',
  classification: 'Primary', complexity: 'Low', relationships: 'child:Y', conflicts: 'DE vs FR' }));
rows.push(mk({ primaryName: 'Zeta Dup', accountId: 'F2', group: 'zeta.com',
  classification: 'Duplicate', complexity: 'Low' }));

// Group G7: Section 2 — High complexity, no rel/conf -> Notes empty
rows.push(mk({ primaryName: 'Eta', accountId: 'G1', group: 'eta.com',
  classification: 'Primary', complexity: 'High' }));
rows.push(mk({ primaryName: 'Eta Dup', accountId: 'G2', group: 'eta.com',
  classification: 'Duplicate', complexity: 'High' }));

// Group G8: Unclassified — all Unclassified, two records
rows.push(mk({ primaryName: 'Theta1', accountId: 'H1', group: 'theta.com',
  validWebsite: 'Y', classification: 'Unclassified', complexity: 'Low' }));
rows.push(mk({ primaryName: 'Theta2', accountId: 'H2', group: 'theta.com',
  validWebsite: 'N', classification: 'Unclassified', complexity: 'Low' }));

// Group G9: Unclassified — single record
rows.push(mk({ primaryName: 'Iota1', accountId: 'I1', group: 'iota.com',
  validWebsite: 'N', classification: 'Unclassified' }));

// Trailing blank line (should be ignored)
rows.push(mk({}));

// --- Round-trip through the writer + parser to also test CSV I/O -----------
var csvText = RP.toCSV(rows);
var reparsed = RP.parseCSV(csvText);
var result = RP.buildReport(reparsed);
var out = result.rows;

console.log('\n=== Stats ===');
console.log(JSON.stringify(result.stats, null, 2));
console.log('\n=== Warnings ===');
console.log(JSON.stringify(result.warnings, null, 2));
console.log('\n=== Output CSV ===\n' + RP.toCSV(out).replace(/^﻿/, ''));

console.log('\n=== Assertions ===');

// Stats
eq(result.stats.section1Count, 3, 'Section 1 has 3 groups');
eq(result.stats.section2Count, 4, 'Section 2 has 4 groups');
eq(result.stats.unclassifiedGroups, 2, '2 unclassified groups');
eq(result.stats.unclassifiedRecordCount, 3, '3 unclassified records');
eq(result.stats.classifiedGroups, 7, '7 classified groups');

// Structure: header rows in the right places
eq(out[0], ['Likely duplicates'], 'Row 0 = section 1 title');
eq(out[1], ['Low complexity, no billing country conflicts, no parent/children'],
  'Row 1 = section 1 description');
eq(out[2][0], 'Primary Account Name', 'Row 2 = base headings');
eq(out[2][8], 'Duplicate Account IDs', 'Base headings end with Duplicate Account IDs');

// Section 1 rows (rows 3,4,5)
var acme = out[3];
eq(acme[0], 'Acme', 'G1 primary name');
eq(acme[1], 'A1', 'G1 account id');
eq(acme[4], 'Yes', 'G1 valid website = Yes (primary Y)');
eq(acme[8], 'A2; A3', 'G1 duplicate ids joined with semicolon');

eq(out[4][4], 'Yes-from dupe', 'G2 valid website = Yes-from dupe');
eq(out[4][8], 'B2', 'G2 duplicate ids');
eq(out[5][4], 'No', 'G3 valid website = No');

// Blank row after section 1
eq(out[6], [], 'Blank row after section 1');

// Section 2 header block
eq(out[7], ['Need attention'], 'Section 2 title');
eq(out[8], ['Flagged, requires approval'], 'Section 2 description');
eq(out[9][9], 'Notes', 'Section 2 headings include Notes as 10th column');

// Section 2 rows + Notes (rows 10..13 = G4,G5,G6,G7)
eq(out[10][0], 'Delta', 'G4 primary name');
eq(out[10][9], 'has relationship', 'G4 note = has relationship');
eq(out[11][9], 'has country conflict', 'G5 note = has country conflict');
eq(out[12][9], 'has relationship & country conflict', 'G6 note = both');
eq(out[13][9], '', 'G7 note empty (only high complexity)');
eq(out[13][6], 'High', 'G7 complexity = High');

// Blank row after section 2
eq(out[14], [], 'Blank row after section 2');

// Section 3
eq(out[15], ['Unclassified'], 'Section 3 title');
eq(out[16][0], 'Account Name', 'Section 3 first heading renamed to Account Name');
eq(out[16].length, 8, 'Section 3 headings have 8 columns (no Duplicate Account IDs)');

// Section 3 records: G8 has 2 records then blank, then G9 record then blank.
eq(out[17][0], 'Theta1', 'Unclassified record 1 name');
eq(out[17][4], 'Yes', 'Unclassified record uses own W (Theta1 Y -> Yes)');
eq(out[18][0], 'Theta2', 'Unclassified record 2 name');
eq(out[18][4], 'No', 'Unclassified record uses own W (Theta2 N -> No)');
eq(out[19], [], 'Blank row separates unclassified groups');
eq(out[20][0], 'Iota1', 'Second unclassified group record');
eq(out[20][4], 'No', 'Iota1 own W = N -> No');
eq(out[21], [], 'Trailing blank row after last unclassified group');

// No warnings expected for this clean dataset
eq(result.warnings.length, 0, 'No warnings for clean dataset');

// --- Parser edge cases -----------------------------------------------------
console.log('\n=== Parser edge cases ===');
var quoted = RP.parseCSV('a,"b,c","line1\nline2","say ""hi"""\r\nx,y,z,w\r\n');
eq(quoted[0], ['a', 'b,c', 'line1\nline2', 'say "hi"'], 'Quoted commas/newlines/escaped quotes');
eq(quoted[1], ['x', 'y', 'z', 'w'], 'Second row parsed');
eq(RP.parseCSV('﻿a,b\n1,2\n').length, 2, 'BOM stripped, trailing newline dropped');

// --- Malformed group routing ----------------------------------------------
console.log('\n=== Malformed group ===');
var rows2 = [header,
  mk({ primaryName: 'NoPrimary1', accountId: 'M1', group: 'mal.com', classification: 'Duplicate' }),
  mk({ primaryName: 'NoPrimary2', accountId: 'M2', group: 'mal.com', classification: 'Duplicate' })
];
var res2 = RP.buildReport(rows2);
assert(res2.warnings.length >= 1, 'Malformed group produces a warning');
eq(res2.stats.unclassifiedRecordCount, 2, 'Malformed group records routed to Unclassified');

console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
