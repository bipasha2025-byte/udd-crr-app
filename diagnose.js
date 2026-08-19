'use strict';

const PizZip = require('pizzip');
const fs = require('fs');
const path = require('path');

const uploadDir = './uploads';
const uploads = fs.readdirSync(uploadDir)
  .map(f => ({ f, t: fs.statSync(path.join(uploadDir, f)).mtime }))
  .sort((a, b) => b.t - a.t);

const crrBuf = fs.readFileSync(path.join(uploadDir, uploads[0].f));
console.log('CRR file:', uploads[0].f);

const zip = new PizZip(crrBuf);
const docXml = zip.file('word/document.xml').asText();

// Count table rows
const trMatches = docXml.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || [];
console.log('Total table rows:', trMatches.length);

const sizes = trMatches.map(r => r.length);
const maxSize = Math.max(...sizes);
const bigRows = sizes.filter(s => s > 50000).length;
console.log('Max row size:', maxSize, 'chars');
console.log('Rows > 50KB:', bigRows);

// Check for nested tables inside rows — these break simple regex
let nestedCount = 0;
for (const row of trMatches) {
  if (row.indexOf('<w:tbl') !== -1) nestedCount++;
}
console.log('Rows containing nested tables:', nestedCount);

// Count cells
const tcMatches = docXml.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || [];
console.log('Total cells:', tcMatches.length);

// Check media files
const mediaFiles = Object.keys(zip.files).filter(function(f) {
  return f.indexOf('word/media') === 0;
});
console.log('Media files:', mediaFiles.length, mediaFiles.slice(0, 3));

// Check if any w:tc patterns are unclosed/broken after our surgery
// Try doing what populateCRR does and check the output XML validity
const { populateCRR } = require('./populator');
const fields = {
  name: 'Christian Khouri (E631475)',
  crrTitle: 'CRR-SAP-ECC-QM-1609-04',
  uddCreationDate: '25-May-2022',
  developmentType: 'Enhancement',
  reviewer: 'Shamik Das (E633074)',
  developerFunction: 'Developer',
  developerName: 'Christian Khouri (E631475)',
};

console.log('\nRunning populateCRR...');
try {
  const out = populateCRR(crrBuf, fields);
  const outZip = new PizZip(out);
  const outDoc = outZip.file('word/document.xml').asText();

  console.log('Output document.xml length:', outDoc.length);
  console.log('Diff chars:', outDoc.length - docXml.length);

  // Check XML is well-formed at a basic level
  const openW = (outDoc.match(/<w:/g) || []).length;
  const closeW = (outDoc.match(/<\/w:/g) || []).length;
  console.log('Open w: tags:', openW, '  Close w: tags:', closeW);
  console.log('Tag balance diff:', openW - closeW);

  // Check for doubled XML declarations (sign of content duplication)
  const xmlDecl = (outDoc.match(/<\?xml/g) || []).length;
  console.log('XML declarations:', xmlDecl, '(should be 1)');

  // Write diagnostics
  fs.writeFileSync('test-samples/DIAG_generated.docx', out);
  console.log('\nDiagnostic file written to test-samples/DIAG_generated.docx');
} catch (e) {
  console.error('populateCRR error:', e.message);
  console.error(e.stack);
}
