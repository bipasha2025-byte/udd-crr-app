'use strict';

/**
 * Regression tests for the two root causes of the "empty Copied Objects table" bug:
 *
 * Root cause 1: APP_COMPONENT_CATEGORIES was missing 'package',
 *               'table index', 'table index (secondary index)', 'secondary index',
 *               and other valid SAP object type labels. When mammoth extracted a UDD
 *               whose section 7.2 used one of these labels as a category, the entire
 *               block was silently skipped → appComponents = [].
 *
 * Root cause 2: looksLikeSAPObject() only accepted names starting with Z, Y, or /.
 *               Standard SAP objects (VLPMA, VBAK, T001, MARA) and secondary-index
 *               notation (VLPMA~ZVL) were rejected → names array empty → block skipped.
 *
 * Reference example from spec:
 *   UDD 7.2: Module/SubModule Area: WM, Development Class: ZUWM,
 *            Secondary Index: VLPMA~ZVL with CRQ175469
 *   Expected CRR rows:
 *     row 1 → Name: ZUWM,      Object Type: Package,                     Comment: (blank)
 *     row 2 → Name: VLPMA~ZVL, Object Type: Table Index (Secondary Index), Comment: CRQ175469
 */

const { extractFieldsFromUDD } = require('./extractor');

let passed = 0;
let failed = 0;
function assert(condition, message) {
  if (condition) { console.log(`  ✓  ${message}`); passed++; }
  else { console.error(`  ✗  ${message}`); failed++; }
}

// ─────────────────────────────────────────────────────────────────
// TEST 1: 'Package' category — previously not in APP_COMPONENT_CATEGORIES
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test 1] Category "Package" — working reference ZUWM example');
{
  const udd = [
    '', 'unit detailed design', '', 'SAP ECC/6.0', '',
    '7.2 System Components', '',
    'App Components & Objects',
    'Name', 'Existing', 'New', 'Upgrade Implications', '',
    'Package',   // <── was missing from APP_COMPONENT_CATEGORIES
    'ZUWM',      // Z custom object — always worked
    'Y',
    '',           // New blank
    '',           // Upgrade blank (no CRQ for Package)
    '',
  ].join('\n');

  const fields = extractFieldsFromUDD(udd);
  const comps = fields.appComponents;

  assert(Array.isArray(comps) && comps.length === 1,
    `1 component extracted (got ${comps ? comps.length : 'null'})`);
  assert(comps && comps[0] && comps[0].name === 'ZUWM',
    `name = "ZUWM" (got "${comps && comps[0] && comps[0].name}")`);
  assert(comps && comps[0] && comps[0].objectType === 'Package',
    `objectType = "Package" (got "${comps && comps[0] && comps[0].objectType}")`);
  assert(comps && comps[0] && comps[0].comment === null,
    `comment = null (no CRQ) (got "${comps && comps[0] && comps[0].comment}")`);
}

// ─────────────────────────────────────────────────────────────────
// TEST 2: 'Table Index (Secondary Index)' category + VLPMA~ZVL object name
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test 2] Category "Table Index (Secondary Index)" + VLPMA~ZVL object name with CRQ');
{
  const udd = [
    '', 'unit detailed design', '', 'SAP ECC/6.0', '',
    '7.2 System Components', '',
    'App Components & Objects',
    'Name', 'Existing', 'New', 'Upgrade Implications', '',
    'Table Index (Secondary Index)',  // <── was missing from APP_COMPONENT_CATEGORIES
    'VLPMA~ZVL',                      // <── had tilde, starts with V — was rejected by looksLikeSAPObject
    'Y',
    '',           // New blank
    'CRQ175469',  // Upgrade CRQ
    '',
  ].join('\n');

  const fields = extractFieldsFromUDD(udd);
  const comps = fields.appComponents;

  assert(Array.isArray(comps) && comps.length === 1,
    `1 component extracted (got ${comps ? comps.length : 'null'})`);
  assert(comps && comps[0] && comps[0].name === 'VLPMA~ZVL',
    `name = "VLPMA~ZVL" (got "${comps && comps[0] && comps[0].name}")`);
  assert(comps && comps[0] && comps[0].objectType === 'Table Index (Secondary Index)',
    `objectType = "Table Index (Secondary Index)" (got "${comps && comps[0] && comps[0].objectType}")`);
  assert(comps && comps[0] && comps[0].comment === 'CRQ175469',
    `comment = "CRQ175469" (got "${comps && comps[0] && comps[0].comment}")`);
}

// ─────────────────────────────────────────────────────────────────
// TEST 3: Full working reference example from spec (both blocks together)
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test 3] Full working reference: Package ZUWM + Table Index VLPMA~ZVL with CRQ');
{
  const udd = [
    '', 'unit detailed design', '', 'SAP ECC/6.0', '',
    '7.2 System Components', '',
    'App Components & Objects',
    'Name', 'Existing', 'New', 'Upgrade Implications', '',
    'Package',
    'ZUWM',
    'Y',
    '',     // New blank
    '',     // Upgrade blank (no CRQ for ZUWM)
    '',
    'Table Index (Secondary Index)',
    'VLPMA~ZVL',
    'Y',
    '',     // New blank
    'CRQ175469',
    '',
  ].join('\n');

  const fields = extractFieldsFromUDD(udd);
  const comps = fields.appComponents;

  assert(Array.isArray(comps) && comps.length === 2,
    `2 components extracted (got ${comps ? comps.length : 'null'})`);
  if (comps && comps.length >= 2) {
    assert(comps[0].name === 'ZUWM' && comps[0].objectType === 'Package' && comps[0].comment === null,
      `row 0: ZUWM / Package / null`);
    assert(comps[1].name === 'VLPMA~ZVL' && comps[1].objectType === 'Table Index (Secondary Index)' && comps[1].comment === 'CRQ175469',
      `row 1: VLPMA~ZVL / Table Index (Secondary Index) / CRQ175469`);
  }
}

// ─────────────────────────────────────────────────────────────────
// TEST 4: Standard SAP table names starting with non-Z/Y letters
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test 4] Standard SAP object names (not Z/Y): VLPMA, VBAK, T001, MARA, BSEG');
{
  const udd = [
    '', 'unit detailed design', '', 'SAP ECC/6.0', '',
    '7.2 System Components', '',
    'App Components & Objects',
    'Name', 'Existing', 'New', 'Upgrade Implications', '',
    'Table',
    'VLPMA',   // starts with V
    'VBAK',    // starts with V
    'T001',    // starts with T
    'MARA',    // starts with M
    'BSEG',    // starts with B
    'Y', 'Y', 'Y', 'Y', 'Y',  // Existing column
    // No New entries
    '', '', '', '', '',        // Upgrade blanks (5 objects, 5 blanks)
    '',
  ].join('\n');

  const fields = extractFieldsFromUDD(udd);
  const comps = fields.appComponents;

  assert(Array.isArray(comps) && comps.length === 5,
    `5 components extracted (got ${comps ? comps.length : 'null'})`);
  if (comps && comps.length >= 5) {
    assert(comps[0].name === 'VLPMA', `VLPMA extracted`);
    assert(comps[1].name === 'VBAK',  `VBAK extracted`);
    assert(comps[2].name === 'T001',  `T001 extracted`);
    assert(comps[3].name === 'MARA',  `MARA extracted`);
    assert(comps[4].name === 'BSEG',  `BSEG extracted`);
  }
}

// ─────────────────────────────────────────────────────────────────
// TEST 5: 'Table Index' (short form) category label
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test 5] Category "Table Index" (short form) with tilde object');
{
  const udd = [
    '', 'unit detailed design', '', 'SAP ECC/6.0', '',
    '7.2 System Components', '',
    'App Components & Objects',
    'Name', 'Existing', 'New', 'Upgrade Implications', '',
    'Table Index',     // short form of the label
    'VLPMA~ZVL',
    'Y',
    '',                // New blank
    'CRQ175469',       // Upgrade CRQ
    '',
  ].join('\n');

  const fields = extractFieldsFromUDD(udd);
  const comps = fields.appComponents;

  assert(Array.isArray(comps) && comps.length === 1,
    `1 component extracted (got ${comps ? comps.length : 'null'})`);
  assert(comps && comps[0] && comps[0].name === 'VLPMA~ZVL', `VLPMA~ZVL extracted`);
  assert(comps && comps[0] && comps[0].comment === 'CRQ175469', `CRQ175469 as comment`);
}

// ─────────────────────────────────────────────────────────────────
// TEST 6: looksLikeSAPObject does NOT accidentally match non-object strings
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test 6] looksLikeSAPObject — does not match column headers / CRQ / Y/N');
{
  const udd = [
    '', 'unit detailed design', '', 'SAP ECC/6.0', '',
    '7.2 System Components', '',
    'App Components & Objects',
    'Name', 'Existing', 'New', 'Upgrade Implications', '',
    'Report',
    'ZQM_WE16_JOB',   // valid SAP object
    'Y',
    '',
    'CRQ10001111',    // must NOT appear as an object name
    '',
  ].join('\n');

  const fields = extractFieldsFromUDD(udd);
  const comps = fields.appComponents;

  assert(Array.isArray(comps) && comps.length === 1,
    `exactly 1 component (CRQ not treated as object name) — got ${comps ? comps.length : 'null'}`);
  assert(comps && comps[0] && comps[0].name === 'ZQM_WE16_JOB',
    `name = "ZQM_WE16_JOB"`);
  assert(comps && comps[0] && comps[0].comment === 'CRQ10001111',
    `CRQ10001111 is in comment, not name`);
}

// ─────────────────────────────────────────────────────────────────
console.log(`\n${'─'.repeat(50)}`);
console.log(`Results: ${passed} passed, ${failed} failed`);
if (failed > 0) {
  console.error('SOME TESTS FAILED');
  process.exit(1);
} else {
  console.log('ALL TESTS PASSED ✓');
  process.exit(0);
}
