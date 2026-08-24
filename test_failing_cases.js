'use strict';

/**
 * Regression tests covering the bug fixes for the "empty/incorrect Copied objects table" defect.
 *
 * Fix 1: APP_COMPONENT_CATEGORIES was missing 'package', 'table index',
 *   'table index (secondary index)', 'secondary index', 'development class', etc.
 *   → block skipped → appComponents = [].
 *
 * Fix 2: looksLikeSAPObject() only accepted Z/Y//-prefixed names.
 *   Standard SAP objects (VLPMA, VBAK, T001) and tilde notation (VLPMA~ZVL)
 *   were rejected → names array empty → block skipped.
 *
 * Fix 3a: 'Development Class' missing from APP_COMPONENT_CATEGORIES → ZUWM row dropped.
 * Fix 3b: objectType must be a direct, exact copy of the UDD label — no translation.
 * Fix 3c: codeVersion: UDD 7.2 has no version column → leave blank (DIRECT_COPY_IF_AVAILABLE).
 *
 * Reference example (spec exact input):
 *   UDD row 1: label='Development Class', name='ZUWM',      existing=Y, new='',  upgrade=''
 *   UDD row 2: label='Secondary Index',   name='VLPMA~ZVL', existing='', new=Y, upgrade=''
 *   Expected CRR:
 *     row 1 → Name: ZUWM,      Object Type: Development Class, Code Version: '', Comment: ''
 *     row 2 → Name: VLPMA~ZVL, Object Type: Secondary Index,   Code Version: '', Comment: ''
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
    `objectType = "Package" (direct copy of UDD label) (got "${comps && comps[0] && comps[0].objectType}")`);
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
    `objectType = "Table Index (Secondary Index)" (direct copy of UDD label) (got "${comps && comps[0] && comps[0].objectType}")`);
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
      `row 0: ZUWM / Package (direct copy) / null`);
    assert(comps[1].name === 'VLPMA~ZVL' && comps[1].objectType === 'Table Index (Secondary Index)' && comps[1].comment === 'CRQ175469',
      `row 1: VLPMA~ZVL / Table Index (Secondary Index) (direct copy) / CRQ175469`);
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
// TEST 7: Exact failing case from bug report — Development Class + Secondary Index
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test 7] Exact bug report case: Development Class/ZUWM + Secondary Index/VLPMA~ZVL');
{
  const udd = [
    '', 'unit detailed design', '', 'SAP ECC/6.0', '',
    '7.2 System Components', '',
    'App Components & Objects',
    'Name', 'Existing', 'New', 'Upgrade Implications', '',
    'Development Class',  // UDD label — copied directly to CRR Object Type
    'ZUWM',
    'Y',  // Existing = Y
    '',   // New = blank
    '',   // Upgrade = blank (no CRQ)
    '',
    'Secondary Index',    // UDD label — copied directly to CRR Object Type
    'VLPMA~ZVL',
    '',   // Existing = blank
    'Y',  // New = Y
    '',   // Upgrade = blank (no CRQ)
    '',
  ].join('\n');

  const fields = extractFieldsFromUDD(udd);
  const c = fields.appComponents;

  assert(Array.isArray(c) && c.length === 2,
    `2 rows generated — ZUWM row not dropped (got ${c ? c.length : 'null'})`);
  if (c && c.length >= 2) {
    // Fix 3a: ZUWM row was being dropped because 'Development Class' wasn't in the category set
    assert(c[0].name === 'ZUWM',
      `row 0 name = "ZUWM"`);
    // Fix 3b: objectType is a direct copy of the UDD label — no translation
    assert(c[0].objectType === 'Development Class',
      `row 0 objectType = "Development Class" (direct copy, not "Package") (got "${c[0].objectType}")`);
    // Fix 3c: codeVersion is blank (UDD 7.2 has no version column)
    assert(c[0].codeVersion === '',
      `row 0 codeVersion = "" (blank — no version in UDD 7.2)`);
    assert(c[0].comment === null,
      `row 0 comment = null (no CRQ)`);

    assert(c[1].name === 'VLPMA~ZVL',
      `row 1 name = "VLPMA~ZVL"`);
    // Fix 3b: objectType is a direct copy of the UDD label — no translation
    assert(c[1].objectType === 'Secondary Index',
      `row 1 objectType = "Secondary Index" (direct copy, not "Table Index (Secondary Index)") (got "${c[1].objectType}")`);
    assert(c[1].codeVersion === '',
      `row 1 codeVersion = "" (blank — no version in UDD 7.2)`);
    assert(c[1].comment === null,
      `row 1 comment = null (no CRQ)`);
  }
}

// ─────────────────────────────────────────────────────────────────
// TEST 8: objectType is a direct copy for all category labels
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test 8] objectType is a direct copy of UDD label for all category types');
{
  const directCopyTests = [
    ['Development Class',            'ZDVCLASS',  'Development Class'],
    ['Secondary Index',              'ZTAB~ZIDX', 'Secondary Index'],
    ['Table Index',                  'ZTAB~ZIDX', 'Table Index'],
    ['Table Index (Secondary Index)','ZTAB~ZIDX', 'Table Index (Secondary Index)'],
    ['Module/SubModule Area',        'ZMM',       'Module/SubModule Area'],
    ['Report',                       'ZRPT_TEST', 'Report'],
    ['Structure',                    'Z1STRUCT',  'Structure'],
    ['Package',                      'ZPKG',      'Package'],
    ['Include Program',              'ZINC_TEST', 'Include Program'],
    ['Transaction Code',             'ZTCODE',    'Transaction Code'],
  ];

  for (const [category, objName, expectedType] of directCopyTests) {
    const udd = [
      '', 'unit detailed design', '', 'SAP ECC/6.0', '',
      '7.2 System Components', '',
      'App Components & Objects',
      'Name', 'Existing', 'New', 'Upgrade Implications', '',
      category,
      objName,
      'Y', '', '', '',
    ].join('\n');

    const fields = extractFieldsFromUDD(udd);
    const c = fields.appComponents;
    const gotType = c && c[0] && c[0].objectType;
    assert(gotType === expectedType,
      `"${category}" → objectType="${gotType}" (expected direct copy "${expectedType}")`);
  }
}

// ─────────────────────────────────────────────────────────────────
// TEST 9: codeVersion is blank (no version column in UDD 7.2)
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test 9] codeVersion is blank for all rows (UDD 7.2 has no version column)');
{
  const udd = [
    '', 'unit detailed design', '', 'SAP ECC/6.0', '',
    '7.2 System Components', '',
    'App Components & Objects',
    'Name', 'Existing', 'New', 'Upgrade Implications', '',
    'Report',
    'ZQM_WE16_JOB',
    'Y', '', 'CRQ10001111', '',
  ].join('\n');

  const fields = extractFieldsFromUDD(udd);
  const c = fields.appComponents;

  assert(Array.isArray(c) && c.length === 1, `1 component (got ${c ? c.length : 'null'})`);
  assert(c && c[0] && c[0].codeVersion === '',
    `codeVersion = "" blank (got "${c && c[0] && c[0].codeVersion}")`);
  // Comment still works correctly
  assert(c && c[0] && c[0].comment === 'CRQ10001111',
    `comment = "CRQ10001111" unchanged`);
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
