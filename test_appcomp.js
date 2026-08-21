'use strict';

const { populateCRR } = require('./populator');
const { extractFieldsFromUDD } = require('./extractor');
const PizZip = require('pizzip');

let passed = 0;
let failed = 0;
function assert(condition, message) {
  if (condition) { console.log(`  ✓  ${message}`); passed++; }
  else { console.error(`  ✗  ${message}`); failed++; }
}

// ─────────────────────────────────────────────────────────────────
// TEST A: extractAppComponents — Structure block CRQ mapping
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test A] extractAppComponents — Structure block with blank HEAD cell');
{
  const udd = [
    '',
    'unit detailed design',
    '',
    'SAP ECC/6.0',
    '',
    '7.2 System Components',
    '',
    'App Components & Objects',
    'Name',
    'Existing',
    'New',
    'Upgrade Implications',
    '',
    'Report',
    'ZR_SOME_REPORT',
    'Y',
    '',
    'CRQ10001111',
    '',
    'Structure',
    'Z1EIREP_GRADE',
    'Z1EIREP_HEAD',
    'Z1EIREP_LINE',
    'Y',
    'Y',
    'Y',
    '',
    '',
    'CRQ10003494',
    '',
    '',
    '',
    'CRQ10003494',
    '',
  ].join('\n');

  const fields = extractFieldsFromUDD(udd);
  const comps = fields.appComponents;

  assert(comps.length === 4, `4 components extracted (got ${comps.length})`);
  assert(comps[0].name === 'ZR_SOME_REPORT' && comps[0].comment === 'CRQ10001111',
    `Report row: ZR_SOME_REPORT / CRQ10001111`);
  assert(comps[1].name === 'Z1EIREP_GRADE' && comps[1].comment === 'CRQ10003494',
    `Structure row 1: Z1EIREP_GRADE / CRQ10003494`);
  assert(comps[2].name === 'Z1EIREP_HEAD' && comps[2].comment === null,
    `Structure row 2: Z1EIREP_HEAD / null (blank cell)`);
  assert(comps[3].name === 'Z1EIREP_LINE' && comps[3].comment === 'CRQ10003494',
    `Structure row 3: Z1EIREP_LINE / CRQ10003494`);
}

// ─────────────────────────────────────────────────────────────────
// TEST B: injectCopiedObjectsTable — replaces old rows, preserves header
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test B] injectCopiedObjectsTable — replaces old rows with new components');
{
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:tbl>
  <w:tr>
    <w:tc><w:p><w:r><w:t>Name of Object</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>Object Type</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>Code Version</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t>OLD_OBJECT1</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>Report</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>OldCRQ</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t>OLD_OBJECT2</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>Structure</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
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

  const buf = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });

  const fields = {
    appComponents: [
      { name: 'Z1EIREP_GRADE', objectType: 'Structure', codeVersion: null, comment: 'CRQ10003494' },
      { name: 'Z1EIREP_HEAD',  objectType: 'Structure', codeVersion: null, comment: null },
      { name: 'Z1EIREP_LINE',  objectType: 'Structure', codeVersion: null, comment: 'CRQ10003494' },
    ],
  };

  try {
    const outBuf = populateCRR(buf, fields);
    const outZip = new PizZip(outBuf);
    const outXml = outZip.file('word/document.xml').asText();

    assert(!outXml.includes('OLD_OBJECT1'), 'OLD_OBJECT1 removed');
    assert(!outXml.includes('OLD_OBJECT2'), 'OLD_OBJECT2 removed');
    assert(!outXml.includes('OldCRQ'),      'OldCRQ removed');
    assert(outXml.includes('Z1EIREP_GRADE'),'Z1EIREP_GRADE inserted');
    assert(outXml.includes('Z1EIREP_HEAD'), 'Z1EIREP_HEAD inserted');
    assert(outXml.includes('Z1EIREP_LINE'), 'Z1EIREP_LINE inserted');
    assert(outXml.includes('CRQ10003494'), 'CRQ10003494 present');
    assert(outXml.includes('Name of Object'), 'Header row preserved');
    assert(outXml.includes('Object Type'),    'Header row Object Type preserved');
  } catch (e) {
    console.error('  ✗  injectCopiedObjectsTable threw:', e.message);
    failed++;
  }
}

// ─────────────────────────────────────────────────────────────────
// TEST C: injectCopiedObjectsTable — no existing data rows (insert after header)
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test C] injectCopiedObjectsTable — no existing data rows');
{
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:tbl>
  <w:tr>
    <w:tc><w:p><w:r><w:t>Name of Object</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>Object Type</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>Code Version</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:tc>
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

  const buf = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });

  const fields = {
    appComponents: [
      { name: 'ZR_REPORT_A', objectType: 'Report', codeVersion: null, comment: 'CRQ99999' },
    ],
  };

  try {
    const outBuf = populateCRR(buf, fields);
    const outZip = new PizZip(outBuf);
    const outXml = outZip.file('word/document.xml').asText();

    assert(outXml.includes('ZR_REPORT_A'), 'ZR_REPORT_A inserted after header');
    assert(outXml.includes('CRQ99999'),    'CRQ99999 inserted');
    assert(outXml.includes('Name of Object'), 'Header row preserved');
  } catch (e) {
    console.error('  ✗  Test C threw:', e.message);
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
