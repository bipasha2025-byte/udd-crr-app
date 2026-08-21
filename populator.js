'use strict';

/**
 * CRR Populator — Opens the CRR DOCX, locates blank fields by their labels,
 * and populates them with extracted values from the UDD.
 *
 * Strategy: operate directly on the OOXML XML via PizZip so that ALL
 * formatting (fonts, borders, tables, headers, footers, etc.) is preserved.
 */

const PizZip = require('pizzip');

// ──────────────────────────────────────────────────────────────────────────────
// XML helpers
// ──────────────────────────────────────────────────────────────────────────────

function decodeXmlEntities(str) {
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n, 10)));
}

function encodeXmlEntities(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Extract all <w:t> text content from an XML string, concatenated.
 */
function extractTextFromXml(xml) {
  const matches = xml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
  return matches.map(m => {
    const inner = m.replace(/<w:t[^>]*>/, '').replace(/<\/w:t>/, '');
    return decodeXmlEntities(inner);
  }).join('');
}

function normalizeLabel(str) {
  return str.toLowerCase().replace(/[:\-_\/\s]+/g, ' ').trim();
}

// ──────────────────────────────────────────────────────────────────────────────
// Field injection specs
// ──────────────────────────────────────────────────────────────────────────────

function buildFieldSpecs(fields) {
  return [
    {
      id: 'name',
      labels: ['name', 'author', 'prepared by', 'created by'],
      value: fields.name,
    },
    {
      id: 'documentTitle',
      labels: ['document title', 'crr document title', 'doc title', 'document name'],
      value: fields.crrTitle,
    },
    {
      id: 'crrCreationDate',
      // CRR section 1.1 "CRR Creation date" → today's date (always later than UDD creation date)
      labels: ['crr creation date'],
      value: fields.crrCreationDate,
    },
    {
      id: 'uddCreationDate',
      // CRR section 1.1 "UDD Creation date" → copied exactly from UDD
      labels: ['udd creation date', 'udd creation', 'creation date', 'document creation date', 'date of creation', 'date created', 'created date'],
      value: fields.uddCreationDate,
    },
    {
      id: 'developmentType',
      labels: ['development type', 'dev type', 'type of development'],
      value: fields.developmentType,
    },
    {
      id: 'reviewer',
      // CRR section 1.2: REVIEWER row has [ROLE][FUNCTION][NAME] — 3 columns
      // We populate FUNCTION (col 1) with reviewerFunction and NAME (col 2) with reviewer name
      labels: ['reviewer'],
      value: fields.reviewer,           // goes into NAME column
      isRoleRow: true,
      roleFunction: fields.reviewerFunction || 'Coordinator',  // goes into FUNCTION column
    },
    {
      id: 'developerName',
      // CRR section 1.2: DEVELOPER row has [ROLE][FUNCTION][NAME] — 3 columns
      // FUNCTION column always says "Developer" (fixed), NAME column = developer name from UDD
      labels: ['developer'],
      value: fields.developerName,      // goes into NAME column
      isRoleRow: true,
      roleFunction: 'Developer',        // always hardcoded — never from UDD
    },
    // ── Section 4: REPOSITORY OBJECTS ─────────────────────────────────────────
    {
      id: 'r3Version',
      // UDD 7.2 System Components → R/3 Version (direct copy, no transformation)
      labels: ['r/3 version', 'r3 version', 'r/3 version:'],
      value: fields.r3Version,
    },
    {
      id: 'sourceSystem',
      // UDD 7.2 System Components → Source System (direct copy)
      labels: ['source system', 'source system:'],
      value: fields.sourceSystem,
    },
    {
      id: 'legacySystem',
      // UDD 7.2 System Components → Legacy System (direct copy)
      labels: ['legacy system', 'legacy system:'],
      value: fields.legacySystem,
    },
    {
      id: 'relatedUDDName',
      // Name/identifier of the uploaded UDD file (injected by server.js)
      // "Related Unit Detailed Design" is the exact CRR Section 4 label for this field
      labels: ['related unit detailed design', 'related udd name', 'related udd', 'udd name', 'udd document'],
      value: fields.relatedUDDName,
    },
    {
      id: 'sopConventions',
      // Always fixed: SOP-0011365
      labels: ['standard or language specific conventions used', 'standard or language specific conventions',
               'conventions used', 'sop conventions', 'language specific conventions'],
      value: 'SOP-0011365',
    },
    {
      id: 'devLanguage',
      // Always fixed: ABAP
      labels: ['development language used', 'development language', 'language used', 'dev language'],
      value: 'ABAP',
    },
    {
      id: 'crqNumber',
      // UDD Appendix 1 Revision Log → latest CRQ number
      labels: ['crq number and project name', 'crq number', 'crq no', 'crq'],
      value: fields.crqNumber && fields.projectName
        ? `${fields.projectName}\n${fields.crqNumber}`
        : (fields.crqNumber || fields.projectName || null),
      isMultiPara: true,  // CRQ cell has multiple paragraphs in CRR
    },
  ];
}

// ──────────────────────────────────────────────────────────────────────────────
// Blank detection
// ──────────────────────────────────────────────────────────────────────────────

function isCellBlank(cellXml) {
  const text = extractTextFromXml(cellXml).trim();
  return text === '' || /^[_\-\s\.]{0,10}$/.test(text);
}

function isRunBlank(runText) {
  const t = runText.trim();
  return t === '' || /^[_\-\s\.]{0,10}$/.test(t);
}

// ──────────────────────────────────────────────────────────────────────────────
// Core XML surgery — inject value into a cell
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Replace all text in a cell with the given value, keeping the cell XML
 * structure (formatting, borders, etc.) intact.
 * Clears all existing <w:t> runs then sets the first one to the value.
 */
function injectValueIntoCell(cellXml, value) {
  const encoded = encodeXmlEntities(value);
  let replaced = false;

  const result = cellXml.replace(/<w:t([^>]*)>([^<]*)<\/w:t>/g, (full, attrs, inner) => {
    if (!replaced) {
      replaced = true;
      const newAttrs = /xml:space/.test(attrs) ? attrs : attrs + ' xml:space="preserve"';
      return `<w:t${newAttrs}>${encoded}</w:t>`;
    }
    return `<w:t${attrs}></w:t>`;
  });

  if (!replaced) {
    // No existing <w:t> — insert a new run INSIDE the last <w:p>...</w:p> in the cell
    // A <w:r> MUST be inside a <w:p> — never directly inside <w:tc>
    if (/<\/w:p>/.test(result)) {
      // Insert before the last </w:p> in the cell
      const lastPClose = result.lastIndexOf('</w:p>');
      return result.substring(0, lastPClose) +
             `<w:r><w:t xml:space="preserve">${encoded}</w:t></w:r>` +
             result.substring(lastPClose);
    }
    // Fallback: wrap in a full paragraph before </w:tc>
    return result.replace(/<\/w:tc>/,
      `<w:p><w:r><w:t xml:space="preserve">${encoded}</w:t></w:r></w:p></w:tc>`);
  }
  return result;
}

/**
 * Inject multi-line value into a cell — one line per existing paragraph.
 * If there are fewer paragraphs than lines, extra lines are appended as new paras.
 * value is a string with '\n'-separated parts.
 */
function injectMultiParaIntoCell(cellXml, value) {
  const parts = value.split('\n').map(s => s.trim()).filter(Boolean);
  if (parts.length <= 1) return injectValueIntoCell(cellXml, parts[0] || value);

  const paraRe = /(<w:p[ >][\s\S]*?<\/w:p>)/g;
  const paras = [];
  let m;
  while ((m = paraRe.exec(cellXml)) !== null) {
    paras.push({ xml: m[1], index: m.index, length: m[1].length });
  }
  if (paras.length === 0) return injectValueIntoCell(cellXml, parts.join(' '));

  // Get pPr from first para for cloning
  const pPrMatch = paras[0].xml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  const pPr = pPrMatch ? pPrMatch[0] : '';

  let newParasXml = '';
  for (let k = 0; k < parts.length; k++) {
    const encoded = encodeXmlEntities(parts[k]);
    if (k < paras.length) {
      // Reuse existing para — replace/insert its text
      let paraXml = paras[k].xml;
      let replaced = false;
      paraXml = paraXml.replace(/<w:t([^>]*)>[^<]*<\/w:t>/, (full, attrs) => {
        replaced = true;
        return `<w:t xml:space="preserve">${encoded}</w:t>`;
      });
      if (!replaced) {
        paraXml = paraXml.replace(/<\/w:p>/, `<w:r><w:t xml:space="preserve">${encoded}</w:t></w:r></w:p>`);
      }
      newParasXml += paraXml;
    } else {
      newParasXml += `<w:p>${pPr}<w:r><w:t xml:space="preserve">${encoded}</w:t></w:r></w:p>`;
    }
  }

  // Replace paras span in cellXml
  const firstStart = paras[0].index;
  const lastEnd = paras[paras.length - 1].index + paras[paras.length - 1].length;
  return cellXml.substring(0, firstStart) + newParasXml + cellXml.substring(lastEnd);
}

// ──────────────────────────────────────────────────────────────────────────────
// Table row processing
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Split a row XML into cell objects [{xml, index, length}].
 * Uses a more robust approach that handles nested XML correctly.
 */
function splitRowIntoCells(rowXml) {
  const cells = [];
  // Match <w:tc> ... </w:tc> — the regex is greedy but we use a non-capturing approach
  // to avoid issues with nested structures.
  const re = /(<w:tc[ >][\s\S]*?<\/w:tc>|<w:tc>[\s\S]*?<\/w:tc>)/g;
  let m;
  while ((m = re.exec(rowXml)) !== null) {
    cells.push({ xml: m[1], index: m.index, length: m[1].length });
  }
  return cells;
}

/**
 * Rebuild a row XML by replacing a specific cell's XML.
 */
function replaceCell(rowXml, cellObj, newCellXml) {
  return rowXml.substring(0, cellObj.index) +
         newCellXml +
         rowXml.substring(cellObj.index + cellObj.length);
}

/**
 * Process a table row for a given field spec.
 * Returns { modified: bool, xml: string }.
 *
 * Supports:
 *  - 2-column rows: [Label] [Value]
 *  - 3-column rows: [Role] [Function] [Name]  (CRR section 1.2 style)
 */
function processTableRow(rowXml, spec, injected) {
  if (injected.has(spec.id)) return { modified: false, xml: rowXml };

  const cells = splitRowIntoCells(rowXml);
  if (cells.length === 0) return { modified: false, xml: rowXml };

  const cellTexts = cells.map(c => extractTextFromXml(c.xml));
  const normTexts = cellTexts.map(t => normalizeLabel(t));

  for (let i = 0; i < cells.length; i++) {
    const normText = normTexts[i];

    const isMatch = spec.labels.some(lbl => {
      const normLbl = normalizeLabel(lbl);
      if (spec.isRoleRow) {
        // For role rows: require exact match AND cell must not be empty
        return normText.length > 0 && normText === normLbl;
      }
      // For non-role rows: require the cell text to equal the label closely
      // Also require minimum length to avoid matching single-char cells like "Y"
      if (normText.length < 3) return false;
      return normText === normLbl ||
             (normText.startsWith(normLbl) && normLbl.length >= 4) ||
             (normLbl.length >= 6 && normText.includes(normLbl));
    });

    if (!isMatch) continue;

    // ── 3-column role row: [ROLE][FUNCTION][NAME]  (CRR section 1.2) ─────────
    // i   = ROLE cell  (REVIEWER / DEVELOPER) — do NOT modify
    // i+1 = FUNCTION cell — inject roleFunction value
    // i+2 = NAME cell     — inject value (the person's name)
    if (spec.isRoleRow && cells.length >= i + 3) {
      let newRowXml = rowXml;

      // Step 1: inject FUNCTION into cell i+1
      if (spec.roleFunction) {
        const fnCell = cells[i + 1];
        newRowXml = replaceCell(newRowXml, fnCell, injectValueIntoCell(fnCell.xml, spec.roleFunction));
      }

      // Step 2: inject NAME into cell i+2
      // Re-parse after step 1 so offsets are correct
      const updatedCells = splitRowIntoCells(newRowXml);
      const nameCell = updatedCells[i + 2];
      if (nameCell) {
        newRowXml = replaceCell(newRowXml, nameCell, injectValueIntoCell(nameCell.xml, spec.value));
      }

      injected.add(spec.id);
      return { modified: true, xml: newRowXml };
    }

    // ── 2-column role row fallback: [ROLE][NAME] ──────────────────────────────
    if (spec.isRoleRow && cells.length === i + 2) {
      const nameCell = cells[i + 1];
      const newRowXml = replaceCell(rowXml, nameCell, injectValueIntoCell(nameCell.xml, spec.value));
      injected.add(spec.id);
      return { modified: true, xml: newRowXml };
    }

    // Choose injector based on isMultiPara flag
    const injector = spec.isMultiPara
      ? (cellXml, val) => injectMultiParaIntoCell(cellXml, val)
      : injectValueIntoCell;

    // ── If a next cell exists, ALWAYS inject into it (not into the label cell)
    //    This handles 2-cell [Label:] [Value] rows correctly even when label has a colon.
    if (i + 1 < cells.length) {
      injected.add(spec.id);
      return { modified: true, xml: replaceCell(rowXml, cells[i + 1], injector(cells[i + 1].xml, spec.value)) };
    }

    // ── Single-cell row: label and value are in the same cell (Label: ___)  ────
    const rawText = cellTexts[i];
    const colonPos = rawText.indexOf(':');
    if (colonPos !== -1) {
      const afterColon = rawText.substring(colonPos + 1).trim();
      if (isRunBlank(afterColon) || afterColon === '') {
        const newCellXml = injectAfterColonInCell(cells[i].xml, spec.value);
        injected.add(spec.id);
        return { modified: true, xml: replaceCell(rowXml, cells[i], newCellXml) };
      }
    }
  }

  return { modified: false, xml: rowXml };
}

/**
 * Inject value after a colon inside a cell, preserving the label formatting.
 */
function injectAfterColonInCell(cellXml, value) {
  const encoded = encodeXmlEntities(value);
  const runPattern = /(<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>)/g;
  const runs = [];
  let m;
  while ((m = runPattern.exec(cellXml)) !== null) {
    runs.push({ xml: m[1], index: m.index, length: m[1].length });
  }

  let colonRunIdx = -1;
  for (let i = 0; i < runs.length; i++) {
    if (extractTextFromXml(runs[i].xml).includes(':')) {
      colonRunIdx = i;
      break;
    }
  }

  if (colonRunIdx === -1) {
    return injectValueIntoCell(cellXml, value);
  }

  // Insert AFTER the colon run but we must stay inside the same <w:p>
  // Find the </w:p> that closes the paragraph containing the colon run
  const afterRun = runs[colonRunIdx].index + runs[colonRunIdx].length;
  const pCloseAfter = cellXml.indexOf('</w:p>', afterRun);
  const insertPos = pCloseAfter !== -1 ? pCloseAfter : afterRun;
  const afterColonRun = `<w:r><w:t xml:space="preserve"> ${encoded}</w:t></w:r>`;
  return cellXml.substring(0, insertPos) + afterColonRun + cellXml.substring(insertPos);
}

// ──────────────────────────────────────────────────────────────────────────────
// Paragraph-based population (outside tables)
// ──────────────────────────────────────────────────────────────────────────────

function processParagraphs(xml, spec, injected) {
  if (injected.has(spec.id)) return xml;
  if (spec.isRoleRow) return xml; // role rows only in tables

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
        xml = xml.substring(0, paragraphs[i].index) + newParaXml + xml.substring(paragraphs[i].index + paragraphs[i].length);
        injected.add(spec.id);
        return xml;
      }
    }

    if (i + 1 < paragraphs.length) {
      const nextText = extractTextFromXml(paragraphs[i + 1].xml).trim();
      if (isRunBlank(nextText)) {
        const encoded = encodeXmlEntities(spec.value);
        const newParaXml = paragraphs[i + 1].xml.replace(
          /<w:t([^>]*)>[^<]*<\/w:t>/,
          `<w:t$1 xml:space="preserve">${encoded}</w:t>`
        );
        xml = xml.substring(0, paragraphs[i + 1].index) + newParaXml + xml.substring(paragraphs[i + 1].index + paragraphs[i + 1].length);
        injected.add(spec.id);
        return xml;
      }
    }
  }

  return xml;
}

// ──────────────────────────────────────────────────────────────────────────────
// Cover page name injection
// ──────────────────────────────────────────────────────────────────────────────

/**
 * The CRR cover page has a table where row 1 contains:
 *   Para 0: "" (blank) ← name goes here
 *   Para 1: "GLIMS INTERFACE"
 *   Para 2: "PROCESS COA DATA REPLY"
 *
 * We find the first <w:tr> whose single cell has a blank first paragraph
 * followed by non-blank paragraphs, and inject the name into that blank para.
 */
function injectNameIntoCoverPage(docXml, name, injected) {
  let done = false;
  // Name may be multi-line (e.g. "GLIMS INTERFACE\nPROCESS COA DATA REPLY")
  const nameParts = name.split('\n').map(s => s.trim()).filter(Boolean);

  // Find the SECOND single-cell row (Row 1 = blank name row, Row 0 = title row)
  let singleCellRowCount = 0;
  const result = docXml.replace(/(<w:tr[ >][\s\S]*?<\/w:tr>)/g, (rowXml) => {
    if (done) return rowXml;

    const cells = splitRowIntoCells(rowXml);
    if (cells.length !== 1) return rowXml; // only merged single-cell rows

    singleCellRowCount++;
    if (singleCellRowCount !== 2) return rowXml; // skip Row 0 (title), target Row 1 (blank name row)

    // Get all paragraphs in the cell
    const paras = [];
    const paraRe = /(<w:p[ >][\s\S]*?<\/w:p>)/g;
    let m;
    while ((m = paraRe.exec(cells[0].xml)) !== null) {
      paras.push({ xml: m[1], index: m.index, length: m[1].length });
    }
    if (paras.length === 0) return rowXml;

    // Verify this row is currently blank (all paras have no text content)
    const cellText = extractTextFromXml(cells[0].xml).trim();
    if (cellText !== '') return rowXml; // already has content — skip

    // Copy paragraph properties (w:pPr) from first para to reuse formatting
    const pPrMatch = paras[0].xml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
    const pPr = pPrMatch ? pPrMatch[0] : '';

    // Build replacement paragraph(s): one per name line
    let newParasXml = '';
    for (let k = 0; k < nameParts.length; k++) {
      const encoded = encodeXmlEntities(nameParts[k]);
      if (k === 0 && paras[0]) {
        // Reuse first existing para (preserves its formatting/properties)
        let paraXml = paras[0].xml;
        const hasRun = /<w:r[ >]/.test(paraXml);
        if (hasRun) {
          let replaced = false;
          paraXml = paraXml.replace(/<w:t([^>]*)>[^<]*<\/w:t>/, (full, attrs) => {
            if (!replaced) {
              replaced = true;
              return `<w:t xml:space="preserve">${encoded}</w:t>`;
            }
            return full;
          });
          if (!replaced) {
            paraXml = paraXml.replace(/<\/w:r>/, `<w:t xml:space="preserve">${encoded}</w:t></w:r>`);
          }
        } else {
          paraXml = paraXml.replace(/<\/w:p>/, `<w:r><w:t xml:space="preserve">${encoded}</w:t></w:r></w:p>`);
        }
        newParasXml += paraXml;
      } else if (k < paras.length) {
        // Reuse subsequent existing para
        let paraXml = paras[k].xml;
        const hasRun = /<w:r[ >]/.test(paraXml);
        if (hasRun) {
          let replaced = false;
          paraXml = paraXml.replace(/<w:t([^>]*)>[^<]*<\/w:t>/, (full, attrs) => {
            if (!replaced) { replaced = true; return `<w:t xml:space="preserve">${encoded}</w:t>`; }
            return full;
          });
          if (!replaced) {
            paraXml = paraXml.replace(/<\/w:r>/, `<w:t xml:space="preserve">${encoded}</w:t></w:r>`);
          }
        } else {
          paraXml = paraXml.replace(/<\/w:p>/, `<w:r><w:t xml:space="preserve">${encoded}</w:t></w:r></w:p>`);
        }
        newParasXml += paraXml;
      } else {
        // Need an extra para — clone first para's pPr and add new run
        newParasXml += `<w:p>${pPr}<w:r><w:t xml:space="preserve">${encoded}</w:t></w:r></w:p>`;
      }
    }

    // Replace all paras in cell: used ones + any leftover blank paras
    let rebuiltCellXml = cells[0].xml;
    // Find the span of all paras in the cell and replace with new paras
    const firstParaStart = paras[0].index;
    const lastPara = paras[paras.length - 1];
    const lastParaEnd = lastPara.index + lastPara.length;
    rebuiltCellXml = rebuiltCellXml.substring(0, firstParaStart) +
                     newParasXml +
                     rebuiltCellXml.substring(lastParaEnd);

    done = true;
    injected.add('name');
    return replaceCell(rowXml, cells[0], rebuiltCellXml);
  });

  return result;
}

// ──────────────────────────────────────────────────────────────────────────────
// Review type injection
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Inject the review type selection into the CRR.
 *
 * The CRR Section 4 has two rows:
 *   ROW A: cell[0]=blank   cell[1]="This is a FULL review"
 *   ROW B: cell[0]="Y"/""  cell[1]="Only CHANGED objects have been reviewed"
 *
 * reviewType = 'full_review'            → put "X" in ROW A cell[0], clear ROW B cell[0]
 * reviewType = 'changed_objects_removed' → put "X" in ROW B cell[0], clear ROW A cell[0]
 *
 * We identify the rows by their label text in cell[1] (case-insensitive).
 * We preserve all formatting — only the text content of cell[0] is changed.
 */
function injectReviewType(docXml, reviewType) {
  if (!reviewType) return docXml;

  const isFull    = reviewType === 'full_review';
  const isChanged = reviewType === 'changed_objects_removed';
  if (!isFull && !isChanged) return docXml;

  // We need to process pairs of matching rows. Use a stateful replace.
  return docXml.replace(/(<w:tr[ >][\s\S]*?<\/w:tr>)/g, (rowXml) => {
    const cells = splitRowIntoCells(rowXml);
    if (cells.length < 2) return rowXml;

    const label = extractTextFromXml(cells[1].xml).replace(/\s+/g, ' ').trim().toLowerCase();

    const isFullRow    = /this is a full review/i.test(label);
    const isChangedRow = /only changed objects/i.test(label) ||
                         /changed objects have been/i.test(label) ||
                         /changed objects have been removed/i.test(label) ||
                         /changed objects have been reviewed/i.test(label);

    if (!isFullRow && !isChangedRow) return rowXml;

    // Determine whether to put X or blank in cell[0]
    const putX = (isFullRow && isFull) || (isChangedRow && isChanged);
    const newCell0 = putX
      ? setCellText(cells[0].xml, 'Y')
      : clearCellText(cells[0].xml);

    return replaceCell(rowXml, cells[0], newCell0);
  });
}

/**
 * Set the text content of a cell to the given value.
 * Preserves all cell formatting (tcPr, pPr, rPr). Only touches <w:t> content.
 */
function setCellText(cellXml, value) {
  const encoded = encodeXmlEntities(value);
  // If a run already exists, replace the first <w:t> text
  if (/<w:r[ >]/.test(cellXml)) {
    let replaced = false;
    const result = cellXml.replace(/<w:t([^>]*)>[^<]*<\/w:t>/, (full, attrs) => {
      if (!replaced) {
        replaced = true;
        return `<w:t xml:space="preserve">${encoded}</w:t>`;
      }
      return full;
    });
    if (replaced) return result;
    // Run exists but no <w:t> — insert before </w:r>
    return cellXml.replace(/<\/w:r>/, `<w:t xml:space="preserve">${encoded}</w:t></w:r>`);
  }
  // No run — insert a new run before </w:p>
  return cellXml.replace(/<\/w:p>/, `<w:r><w:t xml:space="preserve">${encoded}</w:t></w:r></w:p>`);
}

/**
 * Clear the text content of a cell (remove all <w:r> runs from paragraphs,
 * leaving the paragraph and cell structure intact).
 */
function clearCellText(cellXml) {
  // Remove all <w:r>...</w:r> blocks from inside <w:p> elements
  return cellXml.replace(/<w:r[ >][\s\S]*?<\/w:r>/g, '');
}

// ──────────────────────────────────────────────────────────────────────────────
// Main population entry point
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Populate the CRR DOCX with extracted fields.
 * @param {Buffer} crrBuffer - raw bytes of the CRR DOCX file
 * @param {Object} fields    - extracted field values
 * @returns {Buffer}          - modified DOCX bytes (valid DOCX, opens in Word)
 */
function populateCRR(crrBuffer, fields) {
  // Load original zip
  const zip = new PizZip(crrBuffer);
  const specs = buildFieldSpecs(fields);
  const injected = new Set();

  // ── Process main document body ────────────────────────────────────────────
  let docXml = zip.file('word/document.xml').asText();

  // ── Special case: inject name into blank first paragraph of cover page row 1 ─
  // The CRR cover page row 1 has paragraphs: [blank][GLIMS INTERFACE][PROCESS COA...]
  // The name goes into that first blank paragraph.
  if (fields.name) {
    docXml = injectNameIntoCoverPage(docXml, fields.name, injected);
  }

  // Pass 1: table rows — only process until all fields are injected
  // We track row index to avoid matching checklist rows deep in the document
  let rowIndex = 0;
  docXml = docXml.replace(/(<w:tr[ >][\s\S]*?<\/w:tr>)/g, (rowXml) => {
    const currentRow = rowIndex++;
    if (injected.size >= specs.length) return rowXml;
    const hasRoleRowsLeft = specs.some(s => s.isRoleRow && !injected.has(s.id) && s.value);
    const hasNonRoleLeft = specs.some(s => !s.isRoleRow && !injected.has(s.id) && s.value);
    if (currentRow > 20 && !hasRoleRowsLeft && !hasNonRoleLeft) return rowXml;
    if (currentRow > 50) return rowXml;
    for (const spec of specs) {
      if (!spec.value || injected.has(spec.id)) continue;
      const result = processTableRow(rowXml, spec, injected);
      if (result.modified) return result.xml;
    }
    return rowXml;
  });

  // Pass 2: paragraphs (non-table fields)
  for (const spec of specs) {
    if (!spec.value || injected.has(spec.id)) continue;
    docXml = processParagraphs(docXml, spec, injected);
  }

  // ── Review type — inject X into selected row, clear the other row ────────────
  if (fields.reviewType) {
    docXml = injectReviewType(docXml, fields.reviewType);
  }

  zip.file('word/document.xml', docXml);

  // ── Process headers and footers ───────────────────────────────────────────
  const headerFooterFiles = Object.keys(zip.files).filter(f =>
    f.startsWith('word/header') || f.startsWith('word/footer')
  );

  for (const hf of headerFooterFiles) {
    let hfXml = zip.file(hf).asText();
    let changed = false;

    hfXml = hfXml.replace(/(<w:tr[ >][\s\S]*?<\/w:tr>)/g, (rowXml) => {
      for (const spec of specs) {
        if (!spec.value || injected.has(spec.id)) continue;
        const result = processTableRow(rowXml, spec, injected);
        if (result.modified) { changed = true; return result.xml; }
      }
      return rowXml;
    });

    for (const spec of specs) {
      if (!spec.value || injected.has(spec.id)) continue;
      const after = processParagraphs(hfXml, spec, injected);
      if (after !== hfXml) { hfXml = after; changed = true; }
    }

    if (changed) zip.file(hf, hfXml);
  }

  // ── Generate output ───────────────────────────────────────────────────────
  // Critical: [Content_Types].xml and all .rels files MUST be stored
  // without compression (STORE), otherwise Word rejects the file.
  // We achieve this by generating with per-file compression options.
  const output = zip.generate({
    type: 'nodebuffer',
    // Default: compress XML content files
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    // Per-file overrides: structural files must be STORE (no compression)
    platform: 'UNIX',
  });

  // PizZip doesn't support per-file compression in generate() directly.
  // Instead we rebuild the zip correctly by re-reading and re-writing
  // with STORE for structural files.
  return rebuildDocx(zip);
}

/**
 * Rebuild the DOCX zip with correct per-file compression:
 * - [Content_Types].xml → STORE (no compression)
 * - _rels/ and word/_rels/ → STORE
 * - All other files → DEFLATE
 *
 * This is required so that Microsoft Word can open the file.
 */
function rebuildDocx(zip) {
  const outZip = new PizZip();

  // Files that MUST be stored uncompressed per OOXML spec
  const mustStore = (name) =>
    name === '[Content_Types].xml' ||
    name.endsWith('.rels') ||
    name.startsWith('_rels/');

  for (const [name, fileObj] of Object.entries(zip.files)) {
    if (fileObj.dir) {
      // Skip directory entries — PizZip adds them automatically
      continue;
    }
    try {
      const content = fileObj.asNodeBuffer ? fileObj.asNodeBuffer() : fileObj.asBinary();
      const useStore = mustStore(name);
      outZip.file(name, content, {
        binary: true,
        compression: useStore ? 'STORE' : 'DEFLATE',
        compressionOptions: useStore ? {} : { level: 6 },
      });
    } catch (e) {
      // If we can't read as buffer, skip (shouldn't happen)
      console.error(`Warning: could not copy file ${name}:`, e.message);
    }
  }

  return outZip.generate({
    type: 'nodebuffer',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

module.exports = { populateCRR, buildFieldSpecs };
