'use strict';

/**
 * Tests for cover page mapping:
 *   A. extractSystemEnvironment — SAP ECC variant
 *   B. extractSystemEnvironment — SAP GTW variant (non-ECC platform)
 *   C. extractSystemEnvironment — SAP S/4HANA variant
 *   D. extractSystemEnvironment — absent (no SAP line)
 *   E. extractFieldsFromUDD — systemEnvironment present in returned fields
 *   F. populateCRR — injectSystemEnvironmentIntoCoverPage replaces ROW 0 para[1]
 *   G. populateCRR — name injection (ROW 1) still works alongside systemEnvironment
 *   H. populateCRR — CODE REVIEW REPORT label (ROW 0 para[0]) is NOT modified
 */

const { extractFieldsFromUDD } = require('./extractor');
const { populateCRR } = require('./populator');
const PizZip = require('pizzip');

let passed = 0;
let failed = 0;
function assert(condition, message) {
  if (condition) { console.log(`  ✓  ${message}`); passed++; }
  else           { console.error(`  ✗  ${message}`); failed++; }
}

// ── Minimal CRR DOCX builder (cover page ROW 0 + ROW 1 only) ─────────────────
function buildMinimalCRR() {
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:tbl>
  <!-- ROW 0: cover header — single merged cell, 2 paras -->
  <w:tr>
    <w:tc>
      <w:p><w:r><w:t>Code review report</w:t></w:r></w:p>
      <w:p><w:r><w:t xml:space="preserve">SAP ECC 6.0 </w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <!-- ROW 1: application name — single merged cell, blank -->
  <w:tr>
    <w:tc>
      <w:p></w:p>
      <w:p></w:p>
    </w:tc>
  </w:tr>
  <!-- ROW 2: document title (2-cell) -->
  <w:tr>
    <w:tc><w:p><w:r><w:t>Document title</w:t></w:r></w:p></w:tc>
    <w:tc><w:p></w:p></w:tc>
  </w:tr>
</w:tbl>
</w:body>
</w:document>`;

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

// ── Test A: extractSystemEnvironment — SAP ECC ───────────────────────────────
console.log('\n[Test A] extractSystemEnvironment — SAP ECC/6.0 variant');
{
  const udd = ['unit detailed design', '', 'SAP ECC/6.0', '', 'GLIMS INTERFACE'].join('\n');
  const fields = extractFieldsFromUDD(udd);
  assert(fields.systemEnvironment === 'SAP ECC/6.0',
    `systemEnvironment = "SAP ECC/6.0" (got "${fields.systemEnvironment}")`);
}

// ── Test B: extractSystemEnvironment — SAP GTW ───────────────────────────────
console.log('\n[Test B] extractSystemEnvironment — SAP GTW 7.5 (non-ECC platform)');
{
  const udd = ['unit detailed design', '', 'SAP GTW 7.5', '', 'FIORI EPAYMENT APP'].join('\n');
  const fields = extractFieldsFromUDD(udd);
  assert(fields.systemEnvironment === 'SAP GTW 7.5',
    `systemEnvironment = "SAP GTW 7.5" (got "${fields.systemEnvironment}")`);
  // name should still pick up the application name line
  assert(fields.name === 'FIORI EPAYMENT APP',
    `name (application name) = "FIORI EPAYMENT APP" (got "${fields.name}")`);
}

// ── Test C: extractSystemEnvironment — SAP S/4HANA ───────────────────────────
console.log('\n[Test C] extractSystemEnvironment — SAP S/4HANA variant');
{
  const udd = ['user design document', '', 'SAP S/4HANA 2022', '', 'MY APPLICATION'].join('\n');
  const fields = extractFieldsFromUDD(udd);
  assert(fields.systemEnvironment === 'SAP S/4HANA 2022',
    `systemEnvironment = "SAP S/4HANA 2022" (got "${fields.systemEnvironment}")`);
}

// ── Test D: extractSystemEnvironment — no SAP line ────────────────────────────
console.log('\n[Test D] extractSystemEnvironment — absent (no SAP platform line)');
{
  const udd = ['user design document', '', 'My custom system', '', 'APP NAME'].join('\n');
  const fields = extractFieldsFromUDD(udd);
  assert(fields.systemEnvironment === null || fields.systemEnvironment === undefined,
    `systemEnvironment = null when no SAP line found (got "${fields.systemEnvironment}")`);
}

// ── Test E: extractFieldsFromUDD — systemEnvironment in returned object ───────
console.log('\n[Test E] extractFieldsFromUDD — systemEnvironment present in returned fields');
{
  const udd = [
    'unit detailed design', '', 'SAP GTW 7.5', '', 'FIORI EPAYMENT APP', '',
    'UDD Creation date', '01-Jan-2025',
    'Development type', 'Enhancement',
    'Roles and responsibilities', 'FUNCTION', 'NAME', '(plus User ID)',
    'CO-AUTHOR (DEV)  ', '', 'Developer', '', 'Dev Person (E123)',
    'REVIEWED BY (CO)', '', 'Reviewer', '', 'Rev Person (E456)',
    'Appendix 1: Revision Log',
    'DOCUMENT VERSION', 'DATE OF THE CHANGE', 'REASONS OF THE CHANGE', 'Transport',
    '01', '10-Jan-2025', 'MY_PROJ', 'CRQ00000111111', 'DE1KAAAA',
  ].join('\n');
  const fields = extractFieldsFromUDD(udd);
  assert('systemEnvironment' in fields,
    'systemEnvironment key present in extractFieldsFromUDD result');
  assert(fields.systemEnvironment === 'SAP GTW 7.5',
    `systemEnvironment = "SAP GTW 7.5" (got "${fields.systemEnvironment}")`);
}

// ── Test F: populateCRR — ROW 0 para[1] replaced with systemEnvironment ───────
console.log('\n[Test F] populateCRR — ROW 0 para[1] (system environment) replaced correctly');
{
  const buf = buildMinimalCRR();
  const fields = { systemEnvironment: 'SAP GTW 7.5' };

  try {
    const outBuf = populateCRR(buf, fields);
    const outZip = new PizZip(outBuf);
    const outXml = outZip.file('word/document.xml').asText();

    assert(outXml.includes('SAP GTW 7.5'),
      'SAP GTW 7.5 injected into ROW 0 para[1]');
    assert(!outXml.includes('SAP ECC 6.0'),
      'Original "SAP ECC 6.0" replaced (no longer present)');
    // ROW 0 para[0] must be untouched
    assert(outXml.includes('Code review report'),
      '"Code review report" label (ROW 0 para[0]) preserved unchanged');
  } catch (e) {
    console.error('  ✗  Test F threw:', e.message);
    failed++;
  }
}

// ── Test G: both systemEnvironment and name inject correctly together ─────────
console.log('\n[Test G] populateCRR — systemEnvironment + name both inject correctly');
{
  const buf = buildMinimalCRR();
  const fields = {
    systemEnvironment: 'SAP GTW 7.5',
    name: 'FIORI EPAYMENT APP',
  };

  try {
    const outBuf = populateCRR(buf, fields);
    const outZip = new PizZip(outBuf);
    const outXml = outZip.file('word/document.xml').asText();

    assert(outXml.includes('SAP GTW 7.5'),      'systemEnvironment injected');
    assert(!outXml.includes('SAP ECC 6.0'),      'old system env removed');
    assert(outXml.includes('FIORI EPAYMENT APP'),'name (application name) injected into ROW 1');
    assert(outXml.includes('Code review report'),'CODE REVIEW REPORT label untouched');
  } catch (e) {
    console.error('  ✗  Test G threw:', e.message);
    failed++;
  }
}

// ── Test H: CODE REVIEW REPORT / "Code review report" is never overwritten ────
console.log('\n[Test H] populateCRR — "Code review report" label is never modified');
{
  const buf = buildMinimalCRR();
  const fields = { systemEnvironment: 'SAP ECC/6.0', name: 'MY APP' };

  try {
    const outBuf = populateCRR(buf, fields);
    const outZip = new PizZip(outBuf);
    const outXml = outZip.file('word/document.xml').asText();
    assert(outXml.includes('Code review report'),
      '"Code review report" para[0] is preserved verbatim');
    assert(outXml.includes('SAP ECC/6.0'),
      'systemEnvironment "SAP ECC/6.0" injected into para[1]');
  } catch (e) {
    console.error('  ✗  Test H threw:', e.message);
    failed++;
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
