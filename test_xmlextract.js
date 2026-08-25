'use strict';

/**
 * Tests for the XML-direct App Components extraction.
 *
 * Column mapping (UDD → CRR):
 *   col[0] Components/Objects/Object Type → CRR Object Type
 *   col[1] Name                           → CRR Name of Object
 *   col[2] Existing                       (ignored)
 *   col[3] New                            (ignored)
 *   col[4] Upgrade Implications           → CRR Comment
 *   Code Version                          → blank ('')
 *
 * Tests:
 *   A. Basic 5-column table — correct field mapping
 *   B. Multi-paragraph name cell — one CRR row per name
 *   C. Comment (Upgrade Implications) copied verbatim
 *   D. Code Version is always blank ('')
 *   E. Rows with no name but with objectType still produce a row
 *   F. Stops at first non-5-column row (next table / section)
 *   G. Real UDD DOCX (UDD-SAPECC-QM-1609-04.docx) — spot checks
 */

const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const { extractAppComponentsFromXml, extractFieldsFromUDD } = require('./extractor');

let passed = 0;
let failed = 0;
function assert(condition, message) {
  if (condition) { console.log(`  ✓  ${message}`); passed++; }
  else           { console.error(`  ✗  ${message}`); failed++; }
}

// ── Helper: build a minimal UDD DOCX buffer with a custom XML ─────────────────
function buildUDDDocx(docXml) {
  const zip = new PizZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
  zip.file('word/document.xml', docXml);
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// ── Helper: build a 5-cell table row ──────────────────────────────────────────
function row5(c0, c1, c2, c3, c4) {
  const cell = (text, paras) => {
    // paras: array of paragraph strings (each becomes a <w:p>)
    const ps = (paras || [text]).map(t =>
      `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`
    ).join('');
    return `<w:tc><w:p><w:r><w:t/></w:r></w:p>${ps}</w:tc>`;
  };
  // c1 may be an array (multi-para)
  const c1Cell = Array.isArray(c1)
    ? `<w:tc>${c1.map(t => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`).join('')}</w:tc>`
    : `<w:tc><w:p><w:r><w:t xml:space="preserve">${c1}</w:t></w:r></w:p></w:tc>`;
  return `<w:tr>
    <w:tc><w:p><w:r><w:t xml:space="preserve">${c0}</w:t></w:r></w:p></w:tc>
    ${c1Cell}
    <w:tc><w:p><w:r><w:t xml:space="preserve">${c2}</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t xml:space="preserve">${c3}</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t xml:space="preserve">${c4}</w:t></w:r></w:p></w:tc>
  </w:tr>`;
}

function makeDoc(rows) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:tbl>${rows.join('')}</w:tbl></w:body></w:document>`;
}

// ─────────────────────────────────────────────────────────────────
// TEST A: basic mapping — objectType, name, comment
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test A] Basic 5-column table — correct field mapping');
{
  const header = row5('App Components &amp; Objects', 'Name', 'Existing', 'New', 'Upgrade Implications');
  const data1  = row5('Report', 'ZQM_WE16_JOB', 'Y', '', '');
  const data2  = row5('Function Group', 'ZQM_COA_DATA_REPLAY', 'Y', '', 'CRQ10003494');
  const buf = buildUDDDocx(makeDoc([header, data1, data2]));
  const comps = extractAppComponentsFromXml(buf);

  assert(comps.length === 2, `2 rows extracted (got ${comps.length})`);
  if (comps.length >= 2) {
    assert(comps[0].objectType === 'Report',
      `row[0].objectType = "Report" (got "${comps[0].objectType}")`);
    assert(comps[0].name === 'ZQM_WE16_JOB',
      `row[0].name = "ZQM_WE16_JOB" (got "${comps[0].name}")`);
    assert(comps[0].comment === null || comps[0].comment === '',
      `row[0].comment = null/blank (got "${comps[0].comment}")`);
    assert(comps[1].objectType === 'Function Group',
      `row[1].objectType = "Function Group" (got "${comps[1].objectType}")`);
    assert(comps[1].name === 'ZQM_COA_DATA_REPLAY',
      `row[1].name = "ZQM_COA_DATA_REPLAY" (got "${comps[1].name}")`);
    assert(comps[1].comment === 'CRQ10003494',
      `row[1].comment = "CRQ10003494" (got "${comps[1].comment}")`);
  }
}

// ─────────────────────────────────────────────────────────────────
// TEST B: multi-paragraph name cell → one row per name
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test B] Multi-paragraph name cell — one CRR row per name');
{
  const header = row5('App Components &amp; Objects', 'Name', 'Existing', 'New', 'Upgrade Implications');
  const data   = row5('Structure', ['Z1EIREP_GRADE', 'Z1EIREP_HEAD', 'Z1EIREP_LINE'], 'Y', '', 'CRQ10003494');
  const buf = buildUDDDocx(makeDoc([header, data]));
  const comps = extractAppComponentsFromXml(buf);

  assert(comps.length === 3, `3 rows (one per name para) — got ${comps.length}`);
  if (comps.length >= 3) {
    assert(comps[0].name === 'Z1EIREP_GRADE', `row[0].name = "Z1EIREP_GRADE"`);
    assert(comps[1].name === 'Z1EIREP_HEAD',  `row[1].name = "Z1EIREP_HEAD"`);
    assert(comps[2].name === 'Z1EIREP_LINE',  `row[2].name = "Z1EIREP_LINE"`);
    assert(comps[0].objectType === 'Structure', `all rows objectType = "Structure"`);
    assert(comps[2].objectType === 'Structure', `row[2].objectType = "Structure"`);
    // All three share the same comment from col[4]
    assert(comps[0].comment === 'CRQ10003494', `row[0].comment = "CRQ10003494"`);
    assert(comps[1].comment === 'CRQ10003494', `row[1].comment = "CRQ10003494"`);
    assert(comps[2].comment === 'CRQ10003494', `row[2].comment = "CRQ10003494"`);
  }
}

// ─────────────────────────────────────────────────────────────────
// TEST C: comment (Upgrade Implications col[4]) copied verbatim
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test C] Comment (Upgrade Implications) copied verbatim from col[4]');
{
  const header = row5('App Components &amp; Objects', 'Name', 'Existing', 'New', 'Upgrade Implications');
  const data   = row5('Basic type', 'ZQM_COAREPLY01', 'Y', '', 'CRQ000010003494');
  const buf = buildUDDDocx(makeDoc([header, data]));
  const comps = extractAppComponentsFromXml(buf);

  assert(comps.length === 1, `1 row (got ${comps.length})`);
  assert(comps[0] && comps[0].comment === 'CRQ000010003494',
    `comment = "CRQ000010003494" (got "${comps[0] && comps[0].comment}")`);
  // col[2] (Existing=Y) must NOT become the comment
  assert(comps[0] && comps[0].comment !== 'Y',
    'Existing column (Y) is not confused with comment');
}

// ─────────────────────────────────────────────────────────────────
// TEST D: Code Version is always blank
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test D] Code Version is always blank (no version column in UDD)');
{
  const header = row5('App Components &amp; Objects', 'Name', 'Existing', 'New', 'Upgrade Implications');
  const data   = row5('Module / Package', 'ZUQM', 'Y', '', '');
  const buf = buildUDDDocx(makeDoc([header, data]));
  const comps = extractAppComponentsFromXml(buf);

  assert(comps.length === 1, `1 row (got ${comps.length})`);
  assert(comps[0] && comps[0].codeVersion === '',
    `codeVersion = "" blank (got "${comps[0] && comps[0].codeVersion}")`);
}

// ─────────────────────────────────────────────────────────────────
// TEST E: stops at non-5-column row
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test E] Stops at first non-5-column row (next table/section)');
{
  const header  = row5('App Components &amp; Objects', 'Name', 'Existing', 'New', 'Upgrade Implications');
  const data1   = row5('Report', 'ZREPORT_A', 'Y', '', '');
  // A 2-column row (different table) — must not be included
  const other   = `<w:tr><w:tc><w:p><w:r><w:t>Data Description</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>N/A</w:t></w:r></w:p></w:tc></w:tr>`;
  const buf = buildUDDDocx(makeDoc([header, data1, other]));
  const comps = extractAppComponentsFromXml(buf);

  assert(comps.length === 1, `Stops at non-5-col row — got ${comps.length} rows`);
  assert(comps[0].name === 'ZREPORT_A', `only ZREPORT_A extracted`);
}

// ─────────────────────────────────────────────────────────────────
// TEST F: Real UDD DOCX — spot check against known objects
// ─────────────────────────────────────────────────────────────────
const realUDDPath = path.join(__dirname, '..', 'UDD-SAPECC-QM-1609-04.docx');
if (fs.existsSync(realUDDPath)) {
  console.log('\n[Test F] Real UDD DOCX — spot checks');
  const buf = fs.readFileSync(realUDDPath);
  const comps = extractAppComponentsFromXml(buf);

  console.log(`  (extracted ${comps.length} rows)`);
  comps.forEach((c, i) => console.log(`    [${i}] objectType="${c.objectType}" name="${c.name}" comment="${c.comment}"`));

  // Known objects from the real UDD (11 data rows: 1+1+1+3+1+1+3 = 11)
  assert(comps.length === 11, `exactly 11 component rows (got ${comps.length})`);
  assert(!comps.some(c => c.objectType.includes('?')),
    '"Overall Review Status?" section row NOT included');

  const byName = Object.fromEntries(comps.map(c => [c.name, c]));

  // Module / Package — ZUQM
  assert(byName['ZUQM'] !== undefined, `ZUQM extracted`);
  assert(byName['ZUQM'] && byName['ZUQM'].objectType === 'Module / Package',
    `ZUQM objectType = "Module / Package" (got "${byName['ZUQM'] && byName['ZUQM'].objectType}")`);
  assert(byName['ZUQM'] && byName['ZUQM'].codeVersion === '',
    `ZUQM codeVersion = "" blank`);

  // Report — ZQM_WE16_JOB
  assert(byName['ZQM_WE16_JOB'] !== undefined, `ZQM_WE16_JOB extracted`);
  assert(byName['ZQM_WE16_JOB'] && byName['ZQM_WE16_JOB'].objectType === 'Report',
    `ZQM_WE16_JOB objectType = "Report"`);

  // Structure — all three names must be separate rows
  const structs = comps.filter(c => c.objectType === 'Structure');
  assert(structs.length === 3, `3 Structure rows extracted (got ${structs.length})`);
  const structNames = structs.map(c => c.name);
  assert(structNames.includes('Z1EIREP_GRADE'), `Z1EIREP_GRADE extracted as Structure`);
  assert(structNames.includes('Z1EIREP_HEAD'),  `Z1EIREP_HEAD extracted as Structure`);
  assert(structNames.includes('Z1EIREP_LINE'),  `Z1EIREP_LINE extracted as Structure`);

  // CRQ comment on Structure rows
  const gradeRow = structs.find(c => c.name === 'Z1EIREP_GRADE');
  assert(gradeRow && /CRQ/i.test(gradeRow.comment || ''),
    `Z1EIREP_GRADE has CRQ comment (got "${gradeRow && gradeRow.comment}")`);

  // Include — three names
  const includes = comps.filter(c => c.objectType === 'Include');
  assert(includes.length === 3, `3 Include rows extracted (got ${includes.length})`);
} else {
  console.log('\n[Test F] SKIPPED — real UDD not found at', realUDDPath);
}

// ─────────────────────────────────────────────────────────────────
// TEST G: extractFieldsFromUDD with uddBuffer uses XML path
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test G] extractFieldsFromUDD — XML path used when uddBuffer provided');
{
  const header = row5('App Components &amp; Objects', 'Name', 'Existing', 'New', 'Upgrade Implications');
  const data   = row5('Function Module', 'Z_MY_FM', 'Y', '', 'CRQ999888');
  const uddBuf = buildUDDDocx(makeDoc([header, data]));
  // Minimal text (no mammoth-parseable App Components)
  const fields = extractFieldsFromUDD('unit detailed design\nSAP ECC/6.0\nMY APP\n', uddBuf);

  assert(Array.isArray(fields.appComponents), 'appComponents is an array');
  assert(fields.appComponents.length === 1,
    `1 component from XML path (got ${fields.appComponents.length})`);
  if (fields.appComponents.length >= 1) {
    assert(fields.appComponents[0].name === 'Z_MY_FM',
      `name = "Z_MY_FM" (got "${fields.appComponents[0].name}")`);
    assert(fields.appComponents[0].objectType === 'Function Module',
      `objectType = "Function Module" (got "${fields.appComponents[0].objectType}")`);
    assert(fields.appComponents[0].comment === 'CRQ999888',
      `comment = "CRQ999888" (got "${fields.appComponents[0].comment}")`);
    assert(fields.appComponents[0].codeVersion === '',
      `codeVersion = "" (got "${fields.appComponents[0].codeVersion}")`);
  }
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
