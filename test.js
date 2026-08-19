'use strict';

/**
 * Self-contained integration test for extractor.js and populator.js
 * Runs without any test framework — exits with code 0 on pass, 1 on fail.
 */

const { extractFieldsFromUDD, validateExtraction } = require('./extractor');
const { populateCRR } = require('./populator');
const PizZip = require('pizzip');

let passed = 0;
let failed = 0;

function assert(condition, message) {
  if (condition) {
    console.log(`  ✓  ${message}`);
    passed++;
  } else {
    console.error(`  ✗  ${message}`);
    failed++;
  }
}

// ──────────────────────────────────────────────────────────────────────────────
// TEST 1: Extractor — colon format
// ──────────────────────────────────────────────────────────────────────────────
console.log('\n[Test 1] Extractor — colon-separated label: value format');
{
  const sampleText = `
User Design Document

Name: John Smith
Document Title: CRR-SAP-ECC-QM-1609-04
UDD Creation Date: 15.08.2026
Development Type: Enhancement

1.1 Development / Document
Reviewer: Jane Doe
Developer Function: ABAP Developer
Developer Name: John Smith
`;
  const fields = extractFieldsFromUDD(sampleText);
  assert(fields.name === 'John Smith', `name = "${fields.name}"`);
  assert(fields.crrTitle === 'CRR-SAP-ECC-QM-1609-04', `crrTitle = "${fields.crrTitle}"`);
  assert(fields.uddCreationDate === '15.08.2026', `uddCreationDate = "${fields.uddCreationDate}"`);
  assert(fields.developmentType === 'Enhancement', `developmentType = "${fields.developmentType}"`);
  assert(fields.reviewer === 'Jane Doe', `reviewer = "${fields.reviewer}"`);
  assert(fields.developerFunction === 'ABAP Developer', `developerFunction = "${fields.developerFunction}"`);
  assert(fields.developerName === 'John Smith', `developerName = "${fields.developerName}"`);

  const errors = validateExtraction(fields);
  assert(errors.length === 0, `validateExtraction returns no errors (got: ${errors.join(', ') || 'none'})`);
}

// ──────────────────────────────────────────────────────────────────────────────
// TEST 2: Extractor — tab-separated (table-like) format
// ──────────────────────────────────────────────────────────────────────────────
console.log('\n[Test 2] Extractor — tab-separated table format');
{
  const sampleText = `
Name\tAlice Müller
Document Title\tCRR-SAP-MM-2024-01
UDD Creation Date\t01.01.2025
Development Type\tNew Development
Reviewer\tBob Jones
Developer Function\tFunctional Consultant
Developer Name\tAlice Müller
`;
  const fields = extractFieldsFromUDD(sampleText);
  assert(fields.name === 'Alice Müller', `name = "${fields.name}"`);
  assert(fields.crrTitle && fields.crrTitle.startsWith('CRR-'), `crrTitle starts with CRR- = "${fields.crrTitle}"`);
  assert(fields.uddCreationDate === '01.01.2025', `uddCreationDate = "${fields.uddCreationDate}"`);
  assert(fields.developmentType === 'New Development', `developmentType = "${fields.developmentType}"`);
}

// ──────────────────────────────────────────────────────────────────────────────
// TEST 3: Extractor — CRR pattern in free text
// ──────────────────────────────────────────────────────────────────────────────
console.log('\n[Test 3] Extractor — CRR pattern extracted from free text');
{
  const sampleText = `
This document describes the changes for CRR-SAP-FI-999-02 enhancement.
Author: Carol White
Creation Date: 10.05.2024
Development Type: Enhancement
Reviewer: Dave Brown
Developer Function: ABAP Lead
Developer Name: Carol White
`;
  const fields = extractFieldsFromUDD(sampleText);
  assert(fields.crrTitle === 'CRR-SAP-FI-999-02', `crrTitle from free text = "${fields.crrTitle}"`);
}

// ──────────────────────────────────────────────────────────────────────────────
// TEST 4: Extractor — case-insensitive labels
// ──────────────────────────────────────────────────────────────────────────────
console.log('\n[Test 4] Extractor — uppercase labels');
{
  const sampleText = `
NAME: Sarah Connor
DOCUMENT TITLE: CRR-SAP-SD-001-01
UDD CREATION DATE: 22.03.2025
DEVELOPMENT TYPE: Correction
REVIEWER: Mike Lee
DEVELOPER FUNCTION: BASIS Admin
DEVELOPER NAME: Sarah Connor
`;
  const fields = extractFieldsFromUDD(sampleText);
  assert(fields.name === 'Sarah Connor', `name upper = "${fields.name}"`);
  assert(fields.crrTitle === 'CRR-SAP-SD-001-01', `crrTitle upper = "${fields.crrTitle}"`);
  assert(fields.developmentType === 'Correction', `developmentType upper = "${fields.developmentType}"`);
}

// ──────────────────────────────────────────────────────────────────────────────
// TEST 5: Validation — missing fields produce errors
// ──────────────────────────────────────────────────────────────────────────────
console.log('\n[Test 5] Validation — missing fields flagged correctly');
{
  const partialText = `Name: Test User\nDevelopment Type: Enhancement`;
  const fields = extractFieldsFromUDD(partialText);
  const errors = validateExtraction(fields);
  assert(errors.length > 0, `Partial fields trigger ${errors.length} error(s)`);
  // With only Name+DevelopmentType in the text, CRR Title/Date/Reviewer/DevFunction are all missing
  assert(errors.some(e => e.includes('UDD Creation Date')), 'Missing UDD Creation Date reported');
  assert(errors.some(e => e.includes('Reviewer')), 'Missing Reviewer reported');
}

// ──────────────────────────────────────────────────────────────────────────────
// TEST 6: Populator — create a minimal DOCX in memory and inject values
// ──────────────────────────────────────────────────────────────────────────────
console.log('\n[Test 6] Populator — XML surgery on minimal DOCX');
{
  // Minimal valid OOXML structure for testing
  const minimalDocXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:tbl>
  <w:tr>
    <w:tc><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t>Document Title</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t>UDD Creation Date</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t>Development Type</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t>Reviewer</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t>Developer Function</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t>Developer Name</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>
</w:body>
</w:document>`;

  // Build a minimal ZIP (DOCX) in memory
  const zip = new PizZip();
  // Required OOXML boilerplate
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/word/document.xml"
 ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
 Target="word/document.xml"/>
</Relationships>`);
  zip.file('word/document.xml', minimalDocXml);
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
`);

  const buf = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });

  const fields = {
    name: 'John Smith',
    crrTitle: 'CRR-SAP-ECC-QM-1609-04',
    uddCreationDate: '15.08.2026',
    developmentType: 'Enhancement',
    reviewer: 'Jane Doe',
    developerFunction: 'ABAP Developer',
    developerName: 'John Smith',
  };

  let outputBuf;
  try {
    outputBuf = populateCRR(buf, fields);
    assert(Buffer.isBuffer(outputBuf), 'populateCRR returns a Buffer');
    assert(outputBuf.length > 100, `Output buffer non-trivial size: ${outputBuf.length} bytes`);

    // Parse output and check values were injected
    const outZip = new PizZip(outputBuf);
    const outXml = outZip.file('word/document.xml').asText();

    assert(outXml.includes('John Smith'), 'Name injected');
    assert(outXml.includes('CRR-SAP-ECC-QM-1609-04'), 'CRR Title injected');
    assert(outXml.includes('15.08.2026'), 'UDD Creation Date injected');
    assert(outXml.includes('Enhancement'), 'Development Type injected');
    assert(outXml.includes('Jane Doe'), 'Reviewer injected');
    assert(outXml.includes('ABAP Developer'), 'Developer Function injected');
  } catch (e) {
    console.error('  ✗  populateCRR threw:', e.message);
    failed++;
  }
}

// ──────────────────────────────────────────────────────────────────────────────
// Summary
// ──────────────────────────────────────────────────────────────────────────────
console.log(`\n${'─'.repeat(50)}`);
console.log(`Results: ${passed} passed, ${failed} failed`);
if (failed > 0) {
  console.error('SOME TESTS FAILED');
  process.exit(1);
} else {
  console.log('ALL TESTS PASSED ✓');
  process.exit(0);
}
