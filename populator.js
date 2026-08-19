'use strict';

/**
 * CRR Populator — Opens the CRR DOCX, locates blank fields by their labels,
 * and populates them with extracted values from the UDD.
 *
 * Strategy: operate directly on the OOXML XML via PizZip so that ALL
 * formatting (fonts, borders, tables, headers, footers, etc.) is preserved.
 *
 * We perform targeted text surgery: find the run/cell that contains a blank
 * placeholder adjacent to a known label and inject the value there.
 */

const PizZip = require('pizzip');
const fs = require('fs');

// ──────────────────────────────────────────────────────────────────────────────
// XML helpers
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Decode XML entities back to plain text for matching purposes.
 */
function decodeXmlEntities(str) {
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n, 10)));
}

/**
 * Encode characters that must be escaped in XML text nodes.
 */
function encodeXmlEntities(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Extract all <w:t> text content from an XML string.
 * Returns a single concatenated string.
 */
function extractTextFromXml(xml) {
  const matches = xml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
  return matches.map(m => {
    const inner = m.replace(/<w:t[^>]*>/, '').replace(/<\/w:t>/, '');
    return decodeXmlEntities(inner);
  }).join('');
}

/**
 * Normalize a label for fuzzy comparison.
 */
function normalizeLabel(str) {
  return str.toLowerCase().replace(/[:\-_\/\s]+/g, ' ').trim();
}

// ──────────────────────────────────────────────────────────────────────────────
// Field definitions — label patterns → value to inject
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Build the list of field injection specs given the extracted fields.
 * Each spec says: "find a cell/run whose text matches one of these labels,
 * then populate the adjacent blank cell (or the same cell if it has a colon pattern)."
 */
function buildFieldSpecs(fields) {
  return [
    {
      id: 'name',
      labels: ['name', 'author', 'prepared by', 'created by'],
      value: fields.name,
      isFirstPage: true,
    },
    {
      id: 'documentTitle',
      labels: ['document title', 'crr document title', 'title', 'doc title', 'document name'],
      value: fields.crrTitle,
    },
    {
      id: 'uddCreationDate',
      labels: ['udd creation date', 'creation date', 'document creation date', 'date of creation', 'created date', 'date created'],
      value: fields.uddCreationDate,
    },
    {
      id: 'developmentType',
      labels: ['development type', 'dev type', 'type of development', 'type', 'change type'],
      value: fields.developmentType,
    },
    {
      id: 'reviewer',
      labels: ['reviewer', 'reviewed by', 'code reviewer', 'technical reviewer'],
      value: fields.reviewer,
    },
    {
      id: 'developerFunction',
      labels: ['developer function', 'function', 'developer role', 'role', 'designation'],
      value: fields.developerFunction,
    },
    {
      id: 'developerName',
      labels: ['developer name', 'developer', 'developed by', 'programmer'],
      value: fields.developerName,
    },
  ];
}

// ──────────────────────────────────────────────────────────────────────────────
// Blank-cell / placeholder detection
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Determines whether a <w:tc> cell's visible text is "blank" (empty, underscores,
 * dashes, or very short non-meaningful content).
 */
function isCellBlank(cellXml) {
  const text = extractTextFromXml(cellXml).trim();
  return text === '' || /^[_\-\s\.]{0,10}$/.test(text);
}

/**
 * Is the run text a blank placeholder?
 */
function isRunBlank(runText) {
  const t = runText.trim();
  return t === '' || /^[_\-\s\.]{0,10}$/.test(t);
}

// ──────────────────────────────────────────────────────────────────────────────
// Core XML Surgery
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Inject a value into the FIRST <w:t> element of an XML snippet (cell or run),
 * preserving the surrounding XML markup.
 * If the cell has multiple runs, we replace text in the first run and clear
 * subsequent run texts to avoid duplication.
 */
function injectValueIntoCell(cellXml, value) {
  const encoded = encodeXmlEntities(value);

  // Replace the first <w:t ...>...</w:t> with our value
  let replaced = false;
  const result = cellXml.replace(/<w:t([^>]*)>([^<]*)<\/w:t>/g, (full, attrs, inner) => {
    if (!replaced) {
      replaced = true;
      // Ensure xml:space="preserve" so leading/trailing spaces are kept
      const hasSpace = /xml:space/.test(attrs);
      const newAttrs = hasSpace ? attrs : attrs + ' xml:space="preserve"';
      return `<w:t${newAttrs}>${encoded}</w:t>`;
    }
    // Clear subsequent text runs in the same cell
    return `<w:t${attrs}></w:t>`;
  });

  if (!replaced) {
    // No <w:t> found — insert one before </w:tc>
    return result.replace(/<\/w:tc>/, `<w:r><w:t xml:space="preserve">${encoded}</w:t></w:r></w:tc>`);
  }
  return result;
}

// ──────────────────────────────────────────────────────────────────────────────
// Table-based population (most common CRR structure)
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Process a single <w:tr> (table row) XML string.
 * Returns { modified: bool, xml: string }
 *
 * Strategy: for each row, extract cell texts. If a cell text matches a label,
 * try to populate the NEXT cell in the same row (or the same cell after a colon).
 */
function processTableRow(rowXml, spec, injected) {
  if (injected.has(spec.id)) return { modified: false, xml: rowXml };

  // Split row into cells
  const cellPattern = /(<w:tc>[\s\S]*?<\/w:tc>)/g;
  const cells = [];
  let match;
  while ((match = cellPattern.exec(rowXml)) !== null) {
    cells.push({ xml: match[1], index: match.index, length: match[1].length });
  }
  if (cells.length === 0) return { modified: false, xml: rowXml };

  const cellTexts = cells.map(c => extractTextFromXml(c.xml));
  const normTexts = cellTexts.map(t => normalizeLabel(t));

  for (let i = 0; i < cells.length; i++) {
    const normText = normTexts[i];

    // Check if this cell matches one of the labels
    const isMatch = spec.labels.some(lbl => {
      const normLbl = normalizeLabel(lbl);
      return normText === normLbl ||
             normText.includes(normLbl) ||
             normLbl.includes(normText.replace(/[:\s]+$/, ''));
    });

    if (!isMatch) continue;

    // Matched! Now find the target cell:
    // Case A: same cell contains colon — value follows colon in same cell
    const rawText = cellTexts[i];
    const colonPos = rawText.indexOf(':');
    if (colonPos !== -1) {
      const afterColon = rawText.substring(colonPos + 1).trim();
      if (isRunBlank(afterColon) || afterColon === '') {
        // Inject after colon in the same cell
        const newCellXml = injectAfterColonInCell(cells[i].xml, spec.value);
        const newRowXml = rowXml.substring(0, cells[i].index) +
                          newCellXml +
                          rowXml.substring(cells[i].index + cells[i].length);
        injected.add(spec.id);
        return { modified: true, xml: newRowXml };
      }
    }

    // Case B: next cell is blank → populate it
    if (i + 1 < cells.length && isCellBlank(cells[i + 1].xml)) {
      const newCellXml = injectValueIntoCell(cells[i + 1].xml, spec.value);
      // Rebuild row with the modified next cell
      const nextCell = cells[i + 1];
      const newRowXml = rowXml.substring(0, nextCell.index) +
                        newCellXml +
                        rowXml.substring(nextCell.index + nextCell.length);
      injected.add(spec.id);
      return { modified: true, xml: newRowXml };
    }

    // Case C: value cell is TWO cells away (e.g., label | : | value)
    if (i + 2 < cells.length && isCellBlank(cells[i + 2].xml)) {
      const newCellXml = injectValueIntoCell(cells[i + 2].xml, spec.value);
      const targetCell = cells[i + 2];
      const newRowXml = rowXml.substring(0, targetCell.index) +
                        newCellXml +
                        rowXml.substring(targetCell.index + targetCell.length);
      injected.add(spec.id);
      return { modified: true, xml: newRowXml };
    }
  }

  return { modified: false, xml: rowXml };
}

/**
 * In a cell that has "Label: " text, inject value after the colon.
 * We do this by manipulating runs rather than replacing the whole text,
 * so that the label formatting is preserved.
 */
function injectAfterColonInCell(cellXml, value) {
  const encoded = encodeXmlEntities(value);

  // Find the run that contains the colon and inject a new run after it
  // OR if the run text ends after the colon, append to that run
  const runPattern = /(<w:r>[\s\S]*?<\/w:r>)/g;
  const runs = [];
  let m;
  while ((m = runPattern.exec(cellXml)) !== null) {
    runs.push({ xml: m[1], index: m.index, length: m[1].length });
  }

  // Collect all text; find which run holds the colon
  let accumulated = '';
  let colonRunIdx = -1;
  let colonPosInRun = -1;
  for (let i = 0; i < runs.length; i++) {
    const runText = extractTextFromXml(runs[i].xml);
    const colonIdx = runText.indexOf(':');
    if (colonIdx !== -1) {
      colonRunIdx = i;
      colonPosInRun = colonIdx;
      break;
    }
    accumulated += runText;
  }

  if (colonRunIdx === -1) {
    // No colon found; just append value to last run
    return injectValueIntoCell(cellXml, value);
  }

  // Append a new run with the value text after the colon run
  const afterColonRun = `<w:r><w:t xml:space="preserve"> ${encoded}</w:t></w:r>`;
  // Insert after the colon-containing run
  const insertPos = runs[colonRunIdx].index + runs[colonRunIdx].length;
  return cellXml.substring(0, insertPos) + afterColonRun + cellXml.substring(insertPos);
}

// ──────────────────────────────────────────────────────────────────────────────
// Paragraph-based population (for name on first page and other non-table fields)
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Attempt to find and populate a field in paragraph-style text
 * (outside tables). Handles "Label: blank" on same line or label on one
 * paragraph and blank on the next.
 */
function processParagraphs(xml, spec, injected) {
  if (injected.has(spec.id)) return xml;

  // Split into paragraphs
  const paraPattern = /(<w:p[ >][\s\S]*?<\/w:p>)/g;
  const paragraphs = [];
  let m;
  while ((m = paraPattern.exec(xml)) !== null) {
    paragraphs.push({ xml: m[1], index: m.index, length: m[1].length });
  }

  for (let i = 0; i < paragraphs.length; i++) {
    const paraText = extractTextFromXml(paragraphs[i].xml);
    const normText = normalizeLabel(paraText);

    const isMatch = spec.labels.some(lbl => {
      const normLbl = normalizeLabel(lbl);
      return normText === normLbl ||
             normText.startsWith(normLbl) ||
             normText.includes(normLbl + ':') ||
             normText.includes(normLbl + ' ');
    });

    if (!isMatch) continue;

    const colonIdx = paraText.indexOf(':');
    if (colonIdx !== -1) {
      const afterColon = paraText.substring(colonIdx + 1).trim();
      if (isRunBlank(afterColon)) {
        const newParaXml = injectAfterColonInCell(paragraphs[i].xml, spec.value);
        xml = xml.substring(0, paragraphs[i].index) +
              newParaXml +
              xml.substring(paragraphs[i].index + paragraphs[i].length);
        injected.add(spec.id);
        return xml;
      }
    }

    // Next paragraph blank?
    if (i + 1 < paragraphs.length) {
      const nextText = extractTextFromXml(paragraphs[i + 1].xml).trim();
      if (isRunBlank(nextText)) {
        const encoded = encodeXmlEntities(spec.value);
        const newParaXml = paragraphs[i + 1].xml.replace(
          /<w:t([^>]*)>[^<]*<\/w:t>/,
          `<w:t$1 xml:space="preserve">${encoded}</w:t>`
        );
        xml = xml.substring(0, paragraphs[i + 1].index) +
              newParaXml +
              xml.substring(paragraphs[i + 1].index + paragraphs[i + 1].length);
        injected.add(spec.id);
        return xml;
      }
    }
  }

  return xml;
}

// ──────────────────────────────────────────────────────────────────────────────
// Main population entry point
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Populate the CRR DOCX with extracted fields.
 * @param {Buffer} crrBuffer - raw bytes of the CRR DOCX file
 * @param {Object} fields    - extracted field values
 * @returns {Buffer}          - modified DOCX bytes
 */
function populateCRR(crrBuffer, fields) {
  const zip = new PizZip(crrBuffer);
  const specs = buildFieldSpecs(fields);
  const injected = new Set();

  // Process main document body (word/document.xml)
  let docXml = zip.file('word/document.xml').asText();

  // ── Pass 1: Table rows ────────────────────────────────────────────────────
  // Process every <w:tr> row in the document
  docXml = docXml.replace(/(<w:tr[ >][\s\S]*?<\/w:tr>)/g, (rowXml) => {
    for (const spec of specs) {
      if (!spec.value || injected.has(spec.id)) continue;
      const result = processTableRow(rowXml, spec, injected);
      if (result.modified) return result.xml;
    }
    return rowXml;
  });

  // ── Pass 2: Paragraph-based fields (for anything not in a table) ──────────
  for (const spec of specs) {
    if (!spec.value || injected.has(spec.id)) continue;
    docXml = processParagraphs(docXml, spec, injected);
  }

  // ── Pass 3: Headers and footers (for name on first page if in header) ─────
  const headerFiles = Object.keys(zip.files).filter(f =>
    f.startsWith('word/header') || f.startsWith('word/footer')
  );

  for (const hf of headerFiles) {
    let hfXml = zip.file(hf).asText();
    let modified = false;

    for (const spec of specs) {
      if (!spec.value || injected.has(spec.id)) continue;

      // Process table rows in header/footer
      const newHfXml = hfXml.replace(/(<w:tr[ >][\s\S]*?<\/w:tr>)/g, (rowXml) => {
        const result = processTableRow(rowXml, spec, injected);
        if (result.modified) { modified = true; return result.xml; }
        return rowXml;
      });
      if (newHfXml !== hfXml) {
        hfXml = newHfXml;
        modified = true;
      }

      // Process paragraphs in header/footer
      const afterPara = processParagraphs(hfXml, spec, injected);
      if (afterPara !== hfXml) {
        hfXml = afterPara;
        modified = true;
      }
    }

    if (modified) {
      zip.file(hf, hfXml);
    }
  }

  // Write back modified document XML
  zip.file('word/document.xml', docXml);

  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

/**
 * Returns a list of field IDs that could not be injected.
 * Useful for diagnostics (after populateCRR we compare injected set).
 */
function getInjectionReport(fields, injectedSet) {
  const allIds = Object.keys(buildFieldSpecs(fields).reduce((acc, s) => {
    acc[s.id] = true;
    return acc;
  }, {}));

  const missed = [];
  for (const id of allIds) {
    if (!injectedSet.has(id) && fields[idToFieldKey(id)]) {
      missed.push(id);
    }
  }
  return missed;
}

function idToFieldKey(id) {
  const map = {
    name: 'name',
    documentTitle: 'crrTitle',
    uddCreationDate: 'uddCreationDate',
    developmentType: 'developmentType',
    reviewer: 'reviewer',
    developerFunction: 'developerFunction',
    developerName: 'developerName',
  };
  return map[id];
}

module.exports = { populateCRR, buildFieldSpecs };
