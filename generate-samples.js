'use strict';

/**
 * Generates sample UDD.docx and CRR_Template.docx for manual end-to-end testing.
 * Run with:  node generate-samples.js
 */

const PizZip = require('pizzip');
const path = require('path');
const fs = require('fs');

// ── Minimal OOXML helpers ─────────────────────────────────────────────────────

function makeDocx(bodyXml) {
  const zip = new PizZip();

  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>`);

  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
`);

  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
<w:body>${bodyXml}<w:sectPr/></w:body>
</w:document>`);

  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function para(text, bold = false, size = '24') {
  const bTag = bold ? '<w:b/>' : '';
  return `<w:p><w:r><w:rPr>${bTag}<w:sz w:val="${size}"/></w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
}

function heading(text) {
  return `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
    <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>${esc(text)}</w:t></w:r></w:p>`;
}

function tableRow(label, value, valueBlank = false) {
  const valText = valueBlank ? '' : esc(value);
  return `<w:tr>
    <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:tcBorders>
      <w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>
      <w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>
    </w:tcBorders></w:tcPr>
    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${esc(label)}</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/><w:tcBorders>
      <w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>
      <w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>
    </w:tcBorders></w:tcPr>
    <w:p><w:r><w:t xml:space="preserve">${valText}</w:t></w:r></w:p></w:tc>
  </w:tr>`;
}

function esc(s) {
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ── Build UDD ─────────────────────────────────────────────────────────────────

const uddBody = [
  para('USER DESIGN DOCUMENT', true, '32'),
  para(''),

  // Name at the top
  `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
    ${tableRow('Name', 'John Smith')}
    ${tableRow('Document Title', 'CRR-SAP-ECC-QM-1609-04')}
  </w:tbl>`,

  para(''),
  heading('1.1 Development / Document'),
  para(''),

  `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
    ${tableRow('UDD Creation Date', '15.08.2026')}
    ${tableRow('Development Type', 'Enhancement')}
    ${tableRow('Module', 'QM')}
    ${tableRow('System', 'SAP ECC')}
  </w:tbl>`,

  para(''),
  heading('1.2 Roles and Responsibility'),
  para(''),

  `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
    ${tableRow('Reviewer', 'Jane Doe')}
    ${tableRow('Developer Function', 'ABAP Developer')}
    ${tableRow('Developer Name', 'John Smith')}
  </w:tbl>`,

  para(''),
  heading('2. Functional Description'),
  para('This enhancement adds quality management validation to the goods receipt process.'),
  para(''),
  heading('3. Technical Description'),
  para('The implementation uses BAdI MBGR_MODIFY_DOCUMENT to intercept goods receipt.'),
].join('\n');

// ── Build CRR Template (blank fields) ────────────────────────────────────────

const crrBody = [
  para('CODE REVIEW RECORD', true, '32'),
  para(''),

  // Blank name field on first page
  `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
    ${tableRow('Name', '', true)}
    ${tableRow('Review Date', '15.08.2026')}
  </w:tbl>`,

  para(''),
  heading('1.1 Development / Document'),
  para(''),

  `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
    ${tableRow('Document Title', '', true)}
    ${tableRow('UDD Creation Date', '', true)}
    ${tableRow('Development Type', '', true)}
    ${tableRow('Module', 'QM')}
    ${tableRow('System', 'SAP ECC')}
  </w:tbl>`,

  para(''),
  heading('1.2 Roles and Responsibility'),
  para(''),

  `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
    ${tableRow('Reviewer', '', true)}
    ${tableRow('Developer Function', '', true)}
    ${tableRow('Developer Name', '', true)}
  </w:tbl>`,

  para(''),
  heading('2. Review Checklist'),
  para(''),

  `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
    ${tableRow('Code follows naming conventions', 'Yes')}
    ${tableRow('Unit tested', 'Yes')}
    ${tableRow('No hardcoded values', 'Yes')}
    ${tableRow('Performance reviewed', 'Yes')}
    ${tableRow('Error handling implemented', 'Yes')}
  </w:tbl>`,

  para(''),
  heading('3. Comments'),
  para('No issues found. Approved for transport.'),
].join('\n');

// ── Write files ───────────────────────────────────────────────────────────────

const outDir = path.join(__dirname, 'test-samples');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);

const uddPath = path.join(outDir, 'UDD_Sample.docx');
const crrPath = path.join(outDir, 'CRR_Template_Sample.docx');

fs.writeFileSync(uddPath, makeDocx(uddBody));
fs.writeFileSync(crrPath, makeDocx(crrBody));

console.log('Sample documents created:');
console.log('  UDD:', uddPath);
console.log('  CRR:', crrPath);
console.log('');
console.log('Now open http://localhost:3000 and upload these two files.');
