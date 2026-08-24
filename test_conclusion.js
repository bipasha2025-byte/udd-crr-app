'use strict';

/**
 * Tests for Section 19 CONCLUSION — Comments or Recommendations
 *
 * Tests:
 *   A. extractConclusionEntries — basic extraction from revision log
 *   B. extractConclusionEntries — deduplication
 *   C. extractConclusionEntries — entries without CRQ are excluded
 *   D. injectConclusion — populates correct content in CRR via populateCRR
 */

const { extractConclusionEntries, extractFieldsFromUDD } = require('./extractor');
const { populateCRR } = require('./populator');
const PizZip = require('pizzip');

let passed = 0;
let failed = 0;
function assert(condition, message) {
  if (condition) { console.log(`  ✓  ${message}`); passed++; }
  else { console.error(`  ✗  ${message}`); failed++; }
}

// ─────────────────────────────────────────────────────────────────
// TEST A: extractConclusionEntries — basic multi-entry extraction
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test A] extractConclusionEntries — basic multi-entry extraction');
{
  const udd = [
    '', 'unit detailed design', '', 'SAP ECC/6.0', '',
    'Appendix', 'Revision Log',
    'Appendix 1: Revision Log',
    'DOCUMENT VERSION',
    'DATE OF THE CHANGE',
    'REASONS OF THE CHANGE',
    'Transport Request',
    '01',
    '10-Jan-2020',
    'UCB_RUN',
    'INC0000089997 & INC00000088365 Slowness issue',
    'CRQ00000175469 - System performance',
    'DE1K111111',
    '02',
    '15-Mar-2022',
    'SAP PE1: InfleXio (INFL_QM_42) - GLIMS Additional Fields mapping to SAP',
    'CRQ0001003494 - Additional fields',
    'DE1K222222',
  ].join('\n');

  const lines = udd.split('\n');
  const entries = extractConclusionEntries(lines);

  assert(Array.isArray(entries) && entries.length === 2,
    `2 entries extracted (got ${entries ? entries.length : 'null'})`);
  if (entries && entries.length >= 2) {
    assert(entries[0].projectName === 'UCB_RUN',
      `entry[0].projectName = "UCB_RUN" (got "${entries[0].projectName}")`);
    assert(entries[0].crqNumber === 'CRQ00000175469',
      `entry[0].crqNumber = "CRQ00000175469" (got "${entries[0].crqNumber}")`);
    assert(entries[1].projectName === 'SAP PE1: InfleXio (INFL_QM_42) - GLIMS Additional Fields mapping to SAP',
      `entry[1].projectName = SAP PE1... (got "${entries[1].projectName}")`);
    assert(entries[1].crqNumber === 'CRQ0001003494',
      `entry[1].crqNumber = "CRQ0001003494" (got "${entries[1].crqNumber}")`);
  }
}

// ─────────────────────────────────────────────────────────────────
// TEST B: extractConclusionEntries — deduplication
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test B] extractConclusionEntries — deduplication of same project+CRQ');
{
  const udd = [
    '', 'unit detailed design', '', 'SAP ECC/6.0', '',
    'Appendix 1: Revision Log',
    'DOCUMENT VERSION', 'DATE OF THE CHANGE', 'REASONS OF THE CHANGE', 'Transport',
    '01', '10-Jan-2020', 'MY_PROJECT', 'CRQ00000111', 'DE1KAAAAAA',
    '02', '15-Mar-2022', 'MY_PROJECT', 'CRQ00000111', 'DE1KBBBBBB',  // exact duplicate
    '03', '01-Jun-2023', 'OTHER_PROJECT', 'CRQ00000222', 'DE1KCCCCCC',
  ].join('\n');

  const lines = udd.split('\n');
  const entries = extractConclusionEntries(lines);

  assert(Array.isArray(entries) && entries.length === 2,
    `2 unique entries (duplicate removed) (got ${entries ? entries.length : 'null'})`);
  if (entries && entries.length >= 1) {
    assert(entries[0].projectName === 'MY_PROJECT', `entry[0] = MY_PROJECT`);
    assert(entries[entries.length - 1].projectName === 'OTHER_PROJECT', `last entry = OTHER_PROJECT`);
  }
}

// ─────────────────────────────────────────────────────────────────
// TEST C: extractConclusionEntries — entries without CRQ excluded
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test C] extractConclusionEntries — entries without CRQ are excluded');
{
  const udd = [
    '', 'unit detailed design', '', 'SAP ECC/6.0', '',
    'Appendix 1: Revision Log',
    'DOCUMENT VERSION', 'DATE OF THE CHANGE', 'REASONS OF THE CHANGE', 'Transport',
    '01', '10-Jan-2020', 'PROJECT_NO_CRQ', 'Some description without CRQ', 'DE1KAAAAAA',
    '02', '15-Mar-2022', 'PROJECT_WITH_CRQ', 'CRQ00000999 - details', 'DE1KBBBBBB',
  ].join('\n');

  const lines = udd.split('\n');
  const entries = extractConclusionEntries(lines);

  assert(Array.isArray(entries) && entries.length === 1,
    `1 entry (no-CRQ entry excluded) (got ${entries ? entries.length : 'null'})`);
  assert(entries && entries[0] && entries[0].projectName === 'PROJECT_WITH_CRQ',
    `retained entry is PROJECT_WITH_CRQ`);
}

// ─────────────────────────────────────────────────────────────────
// TEST D: populateCRR — injects conclusion into CRR DOCX
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test D] populateCRR — injectConclusion writes correct text into Comments or Recommendations row');
{
  // Minimal CRR with conclusion section
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:tbl>
  <w:tr>
    <w:tc><w:p><w:r><w:t>Comments or Recommendations:</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p/><w:p/><w:p/></w:tc>
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
    conclusionEntries: [
      { projectName: 'SAP PE1: InfleXio (INFL_QM_42) - GLIMS Additional Fields mapping to SAP', crqNumber: 'CRQ0001003494' },
      { projectName: 'UCB_RUN', crqNumber: 'CRQ00000322074' },
    ],
  };

  try {
    const outBuf = populateCRR(buf, fields);
    const outZip = new PizZip(outBuf);
    const outXml = outZip.file('word/document.xml').asText();

    assert(outXml.includes('SAP PE1: InfleXio (INFL_QM_42) - GLIMS Additional Fields mapping to SAP'),
      'Project name 1 injected');
    assert(outXml.includes('CRQ0001003494'), 'CRQ 1 injected');
    assert(outXml.includes('UCB_RUN'), 'Project name 2 injected');
    assert(outXml.includes('CRQ00000322074'), 'CRQ 2 injected');
    assert(outXml.includes('All changes for this CRQ are ok and accepted.'),
      'Fixed sentence injected');
    // COMMENTS block must appear after each acceptance sentence
    assert(outXml.includes('COMMENTS:'), 'COMMENTS: label injected');
    assert(outXml.includes('>N/A<'), 'N/A default comment injected');
    const commentMatches = (outXml.match(/COMMENTS:/g) || []).length;
    assert(commentMatches === 2, `COMMENTS: appears once per entry (2 entries → got ${commentMatches})`);
    // The em dash – (U+2013) should appear between project name and CRQ
    assert(outXml.includes('\u2013'), 'Em dash separator present');
    // Label row must remain untouched
    assert(outXml.includes('Comments or Recommendations:'), 'Label row preserved');
  } catch (e) {
    console.error('  ✗  injectConclusion threw:', e.message);
    failed++;
  }
}

// ─────────────────────────────────────────────────────────────────
// TEST E: full pipeline — extractFieldsFromUDD includes conclusionEntries
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test E] extractFieldsFromUDD — conclusionEntries present in returned fields');
{
  const udd = [
    '', 'unit detailed design', '', 'SAP ECC/6.0', '',
    'UDD Creation date', '25-May-2022',
    'Development type', 'Enhancement',
    'Roles and responsibilities', 'FUNCTION', 'NAME', '(plus User ID)',
    'CO-AUTHOR (DEV)  ', '', '', '', 'Developer', '', 'Dev Person (E123)',
    'REVIEWED BY (CO)', '', '', '', 'Reviewer', '', 'Rev Person (E456)',
    'Appendix 1: Revision Log',
    'DOCUMENT VERSION', 'DATE OF THE CHANGE', 'REASONS OF THE CHANGE', 'Transport',
    '01', '10-Jan-2020', 'MY_PROJ', 'CRQ00000111111', 'DE1KAAAA',
  ].join('\n');

  const fields = extractFieldsFromUDD(udd);
  assert(Array.isArray(fields.conclusionEntries),
    'conclusionEntries is an array');
  assert(fields.conclusionEntries && fields.conclusionEntries.length >= 1,
    `conclusionEntries has at least 1 entry (got ${fields.conclusionEntries && fields.conclusionEntries.length})`);
  if (fields.conclusionEntries && fields.conclusionEntries.length >= 1) {
    assert(fields.conclusionEntries[0].projectName === 'MY_PROJ',
      `conclusionEntries[0].projectName = "MY_PROJ"`);
    assert(fields.conclusionEntries[0].crqNumber === 'CRQ00000111111',
      `conclusionEntries[0].crqNumber = "CRQ00000111111"`);
  }
}

// ─────────────────────────────────────────────────────────────────
// TEST F: injectConclusion — CRQ embedded in project name is stripped (no duplicate)
// ─────────────────────────────────────────────────────────────────
console.log('\n[Test F] injectConclusion — CRQ embedded in project name must not appear twice');
{
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:tbl>
  <w:tr>
    <w:tc><w:p><w:r><w:t>Comments or Recommendations:</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p/><w:p/><w:p/></w:tc>
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

  // Cases where the CRQ is already embedded in the projectName (the bug)
  const fields = {
    conclusionEntries: [
      // "CoA Harmonization CRQ110605"   → should produce "CoA Harmonization – CRQ110605:"
      { projectName: 'CoA Harmonization CRQ110605',     crqNumber: 'CRQ110605' },
      // "Metrology: CRQ000000246287"    → should produce "Metrology – CRQ000000246287:"
      { projectName: 'Metrology: CRQ000000246287',       crqNumber: 'CRQ000000246287' },
      // "UCB_RUN" (no embedded CRQ)     → should produce "UCB_RUN – CRQ000000322074:"
      { projectName: 'UCB_RUN',                          crqNumber: 'CRQ000000322074' },
      // "SAP PE1: InfleXio..." (no CRQ) → preserved verbatim
      { projectName: 'SAP PE1: InfleXio (INFL_QM_42) - GLIMS Additional Fields mapping to SAP',
        crqNumber: 'CRQ000010003494' },
    ],
  };

  try {
    const outBuf = populateCRR(buf, fields);
    const outZip = new PizZip(outBuf);
    const outXml = outZip.file('word/document.xml').asText();

    // ── Case 1: "CoA Harmonization CRQ110605" ─────────────────────
    // CRQ must appear exactly once — not duplicated inside project name
    const coaHeading = 'CoA Harmonization \u2013 CRQ110605:';
    assert(outXml.includes(coaHeading),
      `Case 1: heading = "CoA Harmonization – CRQ110605:" (correct)`);
    assert(!outXml.includes('CRQ110605 \u2013 CRQ110605') && !outXml.includes('CRQ110605\u2013CRQ110605'),
      'Case 1: CRQ110605 does NOT appear twice in heading');

    // ── Case 2: "Metrology: CRQ000000246287" ──────────────────────
    const metHeading = 'Metrology \u2013 CRQ000000246287:';
    assert(outXml.includes(metHeading),
      `Case 2: heading = "Metrology – CRQ000000246287:" (correct)`);
    assert(!outXml.includes('CRQ000000246287 \u2013 CRQ000000246287'),
      'Case 2: CRQ000000246287 does NOT appear twice in heading');

    // ── Case 3: "UCB_RUN" (CRQ not embedded — project name must be untouched)
    assert(outXml.includes('UCB_RUN \u2013 CRQ000000322074:'),
      'Case 3: heading = "UCB_RUN – CRQ000000322074:" (name without embedded CRQ preserved)');

    // ── Case 4: long SAP project name (no CRQ embedded — must be preserved verbatim)
    assert(outXml.includes('SAP PE1: InfleXio (INFL_QM_42) - GLIMS Additional Fields mapping to SAP \u2013 CRQ000010003494:'),
      'Case 4: long project name preserved verbatim with CRQ appended once');

    // All four entries produce the fixed sentence
    const matches = (outXml.match(/All changes for this CRQ are ok and accepted\./g) || []).length;
    assert(matches === 4, `4 fixed sentences present (one per entry) — got ${matches}`);
    // All four entries produce a COMMENTS: block with N/A
    const commentCount = (outXml.match(/COMMENTS:/g) || []).length;
    assert(commentCount === 4, `COMMENTS: appears once per entry (4 entries → got ${commentCount})`);
    const naCount = (outXml.match(/>N\/A</g) || []).length;
    assert(naCount === 4, `N/A appears once per entry (4 entries → got ${naCount})`);

  } catch (e) {
    console.error('  ✗  Test F threw:', e.message);
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
