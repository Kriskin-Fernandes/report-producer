/*
 * Generate a realistic sample export CSV that exercises every section and
 * edge case (complexity "- note" suffix, relationships, conflicts,
 * unclassified groups). Writes test/sample-export.csv.
 *   node test/gen-sample.js
 */
var RP = require('../report.js');
var fs = require('fs');

// A few extra "detail" columns to show off detailed view (arbitrary free
// indices not otherwise used).
var EXTRA = { region: 4, mid: 8, created: 12 };

function mk(o) {
  var r = [];
  for (var i = 0; i < 28; i++) r.push(''); // filler columns stay empty (no header)
  Object.keys(o).forEach(function (k) {
    var idx = RP.COL[k] != null ? RP.COL[k] : EXTRA[k];
    r[idx] = o[k];
  });
  return r;
}

var header = [];
for (var i = 0; i < 28; i++) header.push(''); // unused columns have blank headers
header[RP.COL.accountName] = 'Account Name';
header[RP.COL.accountId] = 'Account ID';
header[RP.COL.type] = 'Type';
header[RP.COL.owner] = 'Owner';
header[RP.COL.group] = 'Stripped Domain';
header[RP.COL.classification] = 'Classification';
header[RP.COL.reason] = 'Reason';
header[RP.COL.complexity] = 'Complexity';
header[RP.COL.relationships] = 'Relationships';
header[RP.COL.conflicts] = 'Country Conflict';
header[EXTRA.region] = 'Region';
header[EXTRA.mid] = 'MID';
header[EXTRA.created] = 'Created';

var rows = [header];

// --- Section 1 (Likely duplicates): classified, Low, no rel, no conflict ---
rows.push(mk({ accountName: 'Acme, Inc.', alias: 'ACME Holdings', accountId: 'ACC-001', region: 'EMEA', mid: 'mid_1001', created: '2024-01-15', type: 'Enterprise', owner: 'Alice Smith', group: 'acme.com', classification: 'Primary', reason: 'Same stripped domain', complexity: 'Low' }));
rows.push(mk({ accountName: 'ACME LTD', alias: 'ACME Trading', accountId: 'ACC-002', type: 'Enterprise', owner: 'Alice Smith', group: 'acme.com', classification: 'Duplicate', complexity: 'Low' }));
rows.push(mk({ accountName: 'Acme Group', accountId: 'ACC-003', type: 'SMB', owner: 'Alice Smith', group: 'acme.com', classification: 'Duplicate', complexity: 'Low' }));

// Low but with a complexity "- note" and no rel/conflict -> still Section 1
rows.push(mk({ accountName: 'Beta Co', accountId: 'ACC-010', type: 'SMB', owner: 'Bob Jones', group: 'beta.io', classification: 'Primary', reason: 'Same domain', complexity: 'Low - clean match' }));
rows.push(mk({ accountName: 'Beta Online', accountId: 'ACC-011', type: 'SMB', owner: 'Bob Jones', group: 'beta.io', classification: 'Duplicate', complexity: 'Low' }));

// --- Section 2 (Needs Attention) ---
// relationship only
rows.push(mk({ accountName: 'Gamma Corp', alias: 'Gamma Global', accountId: 'ACC-020', region: 'NA', mid: 'mid_2003', created: '2023-11-02', type: 'Enterprise', owner: 'Carol Lee', group: 'gamma.net', classification: 'Primary', reason: 'Same domain', complexity: 'Low' }));
rows.push(mk({ accountName: 'Gamma Sub', accountId: 'ACC-021', type: 'Enterprise', owner: 'Carol Lee', group: 'gamma.net', classification: 'Duplicate', complexity: 'Low', relationships: 'parent: Gamma Holdings' }));
// country conflict + complexity "- note"
rows.push(mk({ accountName: 'Delta LLC', accountId: 'ACC-030', type: 'Mid-Market', owner: 'Dave Kim', group: 'delta.org', classification: 'Primary', reason: 'Same domain', complexity: 'Medium - VAT number mismatch', conflicts: 'GB vs US' }));
rows.push(mk({ accountName: 'Delta Intl', accountId: 'ACC-031', type: 'Mid-Market', owner: 'Dave Kim', group: 'delta.org', classification: 'Duplicate', complexity: 'Medium' }));
// relationship + conflict + complexity note, primary is NOT first row
rows.push(mk({ accountName: 'Epsilon EU', accountId: 'ACC-041', type: 'Enterprise', owner: 'Erin Fox', group: 'epsilon.co', classification: 'Duplicate', complexity: 'High' }));
rows.push(mk({ accountName: 'Epsilon', alias: 'Epsilon Worldwide', accountId: 'ACC-040', region: 'EMEA', mid: 'mid_4004', created: '2024-05-20', type: 'Enterprise', owner: 'Erin Fox', group: 'epsilon.co', classification: 'Primary', reason: 'Fuzzy name + domain', complexity: 'High - needs finance sign-off', relationships: 'child: Epsilon Group', conflicts: 'DE vs FR' }));

// --- Section 3 (Unclassified) ---
rows.push(mk({ accountName: 'Zeta One', alias: 'Zeta Prime', accountId: 'ACC-050', region: 'APAC', mid: 'mid_5005', created: '2024-02-10', type: 'SMB', owner: 'Fay Ng', group: 'zeta.com', classification: 'Unclassified', reason: 'Needs manual review', complexity: 'Low', relationships: 'sibling: Zeta Two' }));
rows.push(mk({ accountName: 'Zeta Two', accountId: 'ACC-051', type: 'SMB', owner: 'Fay Ng', group: 'zeta.com', classification: 'Unclassified', complexity: 'Low' }));
rows.push(mk({ accountName: 'Zeta Three', accountId: 'ACC-052', type: 'SMB', owner: 'Fay Ng', group: 'zeta.com', classification: 'Unclassified', complexity: 'Low' }));
rows.push(mk({ accountName: 'Eta Solo', accountId: 'ACC-060', type: 'Enterprise', owner: 'Guy Ash', group: 'eta.dev', classification: 'Unclassified', reason: 'Ambiguous', complexity: 'Medium', conflicts: 'IE vs US' }));

fs.writeFileSync(__dirname + '/sample-export.csv', RP.toCSV(rows));
console.log('Sample written: ' + rows.length + ' rows (incl header)');

// --- Opportunities sample (optional upload) --------------------------------
// Columns: A=Opportunity ID, C=Opportunity owner, D=Owner role, F=Account ID,
// X=Last modified date (MM/DD/YYYY). Links to accounts above by Account ID.
var OPP = { id: 0, owner: 2, role: 3, account: 5, stage: 13, modified: 23 };
function mkOpp(o) {
  var r = [];
  for (var i = 0; i < 24; i++) r.push('');
  Object.keys(o).forEach(function (k) { r[OPP[k]] = o[k]; });
  return r;
}
var oppHeader = [];
for (var oi = 0; oi < 24; oi++) oppHeader.push('');
oppHeader[OPP.id] = 'Opportunity ID';
oppHeader[OPP.owner] = 'Opportunity owner';
oppHeader[OPP.role] = 'Owner role';
oppHeader[OPP.account] = 'Account ID';
oppHeader[OPP.stage] = 'Stage';
oppHeader[OPP.modified] = 'Last Modified Date';

var opps = [oppHeader];
// ACC-001 (Acme, in Likely) -> 3 opportunities, varied dates.
opps.push(mkOpp({ id: '006Vk00000VshhF', owner: 'Alice Smith', role: 'Account Executive', account: 'ACC-001', stage: 'Closed Won', modified: '07/02/2026' }));
opps.push(mkOpp({ id: '006Vk00000Vshh1', owner: 'Bruno Diaz', role: 'SDR', account: 'ACC-001', stage: 'Negotiation', modified: '06/18/2026' }));
opps.push(mkOpp({ id: '006Vk00000Vshh2', owner: 'Alice Smith', role: 'Account Executive', account: 'ACC-001', stage: 'Prospecting', modified: '01/10/2026' }));
// ACC-040 (Epsilon, Problematic) -> 1 opportunity.
opps.push(mkOpp({ id: '006Vk00000Vshh9', owner: 'Erin Fox', role: 'Account Manager', account: 'ACC-040', stage: 'Qualification', modified: '05/30/2026' }));
// ACC-050 (Zeta, Unclassified) -> 2 opportunities.
opps.push(mkOpp({ id: '006Vk00000VshhA', owner: 'Fay Ng', role: 'SDR', account: 'ACC-050', stage: 'Closed Lost', modified: '07/14/2026' }));
opps.push(mkOpp({ id: '006Vk00000VshhB', owner: 'Fay Ng', role: 'Account Executive', account: 'ACC-050', stage: 'Proposal', modified: '03/01/2026' }));
// An opportunity that maps to NO account in the report (should be ignored).
opps.push(mkOpp({ id: '006Vk00000VshhZ', owner: 'Nobody', role: 'SDR', account: 'ACC-999', stage: 'Prospecting', modified: '07/01/2026' }));

fs.writeFileSync(__dirname + '/sample-opps.csv', RP.toCSV(opps));
console.log('Opps sample written: ' + opps.length + ' rows (incl header)');
