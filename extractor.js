'use strict';

/**
 * UDD Extractor — intelligently extracts required fields from a UDD document.
 * Works on the raw text content extracted by mammoth.
 */

/**
 * Normalize a label string for comparison:
 * lowercase, collapse whitespace, remove trailing colon/punct
 */
function normalizeLabel(str) {
  return str.toLowerCase().replace(/[:\-_]+$/, '').replace(/\s+/g, ' ').trim();
}

/**
 * Given a list of candidate label strings, find the value that follows
 * the label in the provided lines array.
 * Searches each line for "Label: value" or "Label\n value" patterns.
 */
function findValueByLabels(lines, candidateLabels) {
  const normCandidates = candidateLabels.map(normalizeLabel);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const lineTrimmed = line.trim();

    // Try "Label: value" on the same line
    for (const candidate of normCandidates) {
      const normLine = normalizeLabel(lineTrimmed);
      // Check if line starts with or contains the candidate label followed by a colon/space
      const colonIdx = lineTrimmed.indexOf(':');
      if (colonIdx !== -1) {
        const labelPart = normalizeLabel(lineTrimmed.substring(0, colonIdx));
        if (labelPart === candidate || labelPart.endsWith(candidate) || candidate.endsWith(labelPart)) {
          const value = lineTrimmed.substring(colonIdx + 1).trim();
          if (value && value.length > 0 && !isPlaceholder(value)) {
            return value;
          }
          // value may be on next line(s)
          for (let j = i + 1; j < Math.min(i + 4, lines.length); j++) {
            const nextLine = lines[j].trim();
            if (nextLine && !isPlaceholder(nextLine) && !looksLikeLabel(nextLine)) {
              return nextLine;
            }
          }
        }
      }

      // Also check if entire line (normalized) equals candidate label
      if (normLine === candidate) {
        // Value on next non-empty line
        for (let j = i + 1; j < Math.min(i + 4, lines.length); j++) {
          const nextLine = lines[j].trim();
          if (nextLine && !isPlaceholder(nextLine) && !looksLikeLabel(nextLine)) {
            return nextLine;
          }
        }
      }

      // Check if line contains the label as a standalone cell (table-like, tab-separated)
      const parts = lineTrimmed.split(/\t+/);
      if (parts.length >= 2) {
        for (let p = 0; p < parts.length - 1; p++) {
          const partNorm = normalizeLabel(parts[p]);
          if (partNorm === candidate || partNorm.includes(candidate)) {
            const val = parts[p + 1].trim();
            if (val && !isPlaceholder(val)) return val;
          }
        }
      }
    }
  }
  return null;
}

function isPlaceholder(str) {
  return /^[_\-\s\.]*$/.test(str) || str === '' || /^(tbd|n\/a|none)$/i.test(str.trim());
}

function looksLikeLabel(str) {
  // If it ends with a colon or looks like "Label:" pattern
  return /:\s*$/.test(str) || /^[0-9]+\.[0-9]*/.test(str);
}

/**
 * Extract the Name from the UDD.
 * The name typically appears at the top of the document near "Name:" label
 * or in the document properties / header table.
 */
function extractName(lines) {
  const candidates = ['name', 'author', 'prepared by', 'created by', 'document owner'];
  return findValueByLabels(lines, candidates);
}

/**
 * Extract the CRR Document Title from the UDD.
 * Looks for a value matching pattern: CRR-... (e.g. CRR-SAP-ECC-QM-1609-04)
 */
function extractCRRTitle(lines) {
  // First: scan every line for a CRR-... pattern
  const crrPattern = /\bCRR[-_][A-Z0-9][-A-Z0-9_]{2,}\b/i;
  for (const line of lines) {
    const match = line.match(crrPattern);
    if (match) {
      return match[0].trim();
    }
  }

  // Fallback: look for "document title" label (NOT generic "title" alone — too ambiguous)
  const candidates = ['document title', 'crr document title', 'crr title', 'document name', 'doc title'];
  const val = findValueByLabels(lines, candidates);
  if (val && crrPattern.test(val)) return val.match(crrPattern)[0];
  // Only return the fallback value if it actually looks like a CRR title
  if (val && /CRR/i.test(val)) return val;
  return null;
}

/**
 * Extract UDD Creation Date from UDD section 1.1
 */
function extractUDDCreationDate(lines) {
  const candidates = [
    'udd creation date',
    'creation date',
    'document creation date',
    'date of creation',
    'created date',
    'date created',
    'date',
  ];
  return findValueByLabels(lines, candidates);
}

/**
 * Extract Development Type from UDD section 1.1
 */
function extractDevelopmentType(lines) {
  const candidates = [
    'development type',
    'dev type',
    'type of development',
    'type',
    'change type',
    'request type',
  ];
  return findValueByLabels(lines, candidates);
}

/**
 * Extract Reviewer from UDD section 1.2
 */
function extractReviewer(lines) {
  const candidates = [
    'reviewer',
    'reviewed by',
    'review by',
    'code reviewer',
    'technical reviewer',
  ];
  return findValueByLabels(lines, candidates);
}

/**
 * Extract Developer Function from UDD
 */
function extractDeveloperFunction(lines) {
  const candidates = [
    'function',
    'developer function',
    'developer role',
    'role',
    'designation',
    'job title',
    'position',
  ];

  // Be careful not to confuse with reviewer section
  // Look specifically near "developer" context
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim().toLowerCase();
    if (line.includes('developer') || line.includes('dev function') || line.includes('developer function')) {
      const colonIdx = lines[i].indexOf(':');
      if (colonIdx !== -1) {
        const val = lines[i].substring(colonIdx + 1).trim();
        if (val && !isPlaceholder(val)) return val;
      }
      // check tab-separated
      const parts = lines[i].split(/\t+/);
      if (parts.length >= 2) {
        const normFirst = normalizeLabel(parts[0]);
        if (normFirst.includes('function') || normFirst.includes('role')) {
          const val = parts[1].trim();
          if (val && !isPlaceholder(val)) return val;
        }
      }
    }
  }

  return findValueByLabels(lines, candidates);
}

/**
 * Extract Developer Name from UDD
 */
function extractDeveloperName(lines) {
  const candidates = [
    'developer name',
    'developer',
    'developed by',
    'programmer',
    'abap developer',
  ];

  // Look near rows that explicitly say "developer" + "name"
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    const lineLow = line.toLowerCase();
    if (lineLow.includes('developer') && lineLow.includes('name')) {
      const colonIdx = line.indexOf(':');
      if (colonIdx !== -1) {
        const val = line.substring(colonIdx + 1).trim();
        if (val && !isPlaceholder(val)) return val;
      }
      const parts = line.split(/\t+/);
      if (parts.length >= 2) {
        const val = parts[1].trim();
        if (val && !isPlaceholder(val)) return val;
      }
      // value on next line
      for (let j = i + 1; j < Math.min(i + 3, lines.length); j++) {
        const next = lines[j].trim();
        if (next && !isPlaceholder(next) && !looksLikeLabel(next)) return next;
      }
    }
  }

  return findValueByLabels(lines, candidates);
}

/**
 * Main extraction function.
 * @param {string} rawText - plain text extracted from UDD by mammoth
 * @returns {Object} extracted fields with null for missing ones
 */
function extractFieldsFromUDD(rawText) {
  const lines = rawText.split('\n');

  const name = extractName(lines);
  const crrTitle = extractCRRTitle(lines);
  const uddCreationDate = extractUDDCreationDate(lines);
  const developmentType = extractDevelopmentType(lines);
  const reviewer = extractReviewer(lines);
  const developerFunction = extractDeveloperFunction(lines);
  const developerName = extractDeveloperName(lines);

  return {
    name,
    crrTitle,
    uddCreationDate,
    developmentType,
    reviewer,
    developerFunction,
    developerName,
  };
}

/**
 * Validate extraction results.
 * Returns an array of error strings (empty = all good).
 */
function validateExtraction(fields) {
  const errors = [];
  const fieldNames = {
    name: 'Name',
    crrTitle: 'CRR Document Title',
    uddCreationDate: 'UDD Creation Date',
    developmentType: 'Development Type',
    reviewer: 'Reviewer',
    developerFunction: 'Developer Function',
    developerName: 'Developer Name',
  };

  for (const [key, label] of Object.entries(fieldNames)) {
    if (!fields[key]) {
      errors.push(`"${label}" could not be identified in the uploaded UDD.`);
    }
  }
  return errors;
}

module.exports = { extractFieldsFromUDD, validateExtraction };
