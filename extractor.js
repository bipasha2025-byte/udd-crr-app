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
 * Returns true if a string looks like a real person name:
 * - Contains at least one letter
 * - Not purely numeric
 * - Length > 2
 * - May contain letters, spaces, hyphens, parentheses (for Employee IDs like "John Smith (E123)")
 */
function looksLikeName(str) {
  if (!str || str.trim().length < 2) return false;
  const t = str.trim();
  // Reject pure numbers or very short non-name strings
  if (/^\d+$/.test(t)) return false;
  // Must contain at least one letter
  if (!/[a-zA-Z]/.test(t)) return false;
  // Reject things that look like section numbers or codes
  if (/^[\d\.\-]+$/.test(t)) return false;
  return true;
}

/**
 * Returns true if a string looks like a valid function/role value
 * (not a bare number, not a placeholder)
 */
function looksLikeFunctionValue(str) {
  if (!str || str.trim().length < 2) return false;
  const t = str.trim();
  if (/^\d+$/.test(t)) return false;
  if (!/[a-zA-Z]/.test(t)) return false;
  return true;
}

/**
 * Given a list of candidate label strings, find the value that follows
 * the label in the provided lines array.
 */
function findValueByLabels(lines, candidateLabels, validator) {
  const normCandidates = candidateLabels.map(normalizeLabel);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const lineTrimmed = line.trim();

    for (const candidate of normCandidates) {
      const normLine = normalizeLabel(lineTrimmed);

      // Try "Label: value" on the same line
      const colonIdx = lineTrimmed.indexOf(':');
      if (colonIdx !== -1) {
        const labelPart = normalizeLabel(lineTrimmed.substring(0, colonIdx));
        if (labelPart === candidate || labelPart.endsWith(candidate) || candidate.endsWith(labelPart)) {
          const value = lineTrimmed.substring(colonIdx + 1).trim();
          if (value && value.length > 0 && !isPlaceholder(value)) {
            if (!validator || validator(value)) return value;
          }
          // value may be on next line(s)
          for (let j = i + 1; j < Math.min(i + 4, lines.length); j++) {
            const nextLine = lines[j].trim();
            if (nextLine && !isPlaceholder(nextLine) && !looksLikeLabel(nextLine)) {
              if (!validator || validator(nextLine)) return nextLine;
            }
          }
        }
      }

      // Also check if entire line (normalized) equals candidate label
      if (normLine === candidate) {
        for (let j = i + 1; j < Math.min(i + 4, lines.length); j++) {
          const nextLine = lines[j].trim();
          if (nextLine && !isPlaceholder(nextLine) && !looksLikeLabel(nextLine)) {
            if (!validator || validator(nextLine)) return nextLine;
          }
        }
      }

      // Check tab-separated (table-like)
      const parts = lineTrimmed.split(/\t+/);
      if (parts.length >= 2) {
        for (let p = 0; p < parts.length - 1; p++) {
          const partNorm = normalizeLabel(parts[p]);
          if (partNorm === candidate || partNorm.includes(candidate)) {
            const val = parts[p + 1].trim();
            if (val && !isPlaceholder(val)) {
              if (!validator || validator(val)) return val;
            }
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
  return /:\s*$/.test(str) || /^[0-9]+\.[0-9]*/.test(str);
}

// ──────────────────────────────────────────────────────────────────────────────
// Field extractors
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Extract Name — the author/developer name for the CRR first page.
 * In UDD the name comes from:
 *   - "CO-AUTHOR (DEV)" row (3-col table) → NAME column  (developer name)
 *   - "AUTHOR (FC)" row → NAME column (functional consultant — fallback)
 *   - Generic "name", "author", "prepared by" labels
 */
function extractName(lines) {
  // The "name" field = the project/document name from the UDD cover page.
  // UDD cover structure (mammoth line-by-line):
  //   "unit detailed design"   OR  "user design document"
  //   "SAP ECC/6.0"            ← SAP platform line
  //   "GLIMS INTERFACE"        ← project name line 1   ← we want this
  //   "PROCESS COA DATA REPLY" ← project name line 2   ← and this
  //
  // Strategy 1: find "SAP ECC" line near the top, collect 1-2 non-empty lines after it
  const sapEccIdx = findSapEccLine(lines);
  if (sapEccIdx !== -1) {
    const nameParts = [];
    for (let j = sapEccIdx + 1; j < Math.min(sapEccIdx + 10, lines.length); j++) {
      const t = lines[j].trim();
      if (!t) continue;
      // Stop when we hit a TOC entry or section header (has tab + number)
      if (/\t\d+$/.test(t) || /^table of contents/i.test(t) || /^appendix/i.test(t)) break;
      // Stop when we hit a known section keyword
      if (/^(formal details|development|roles and|section|chapter|\d+\.)/i.test(t)) break;
      nameParts.push(t);
      if (nameParts.length >= 2) break; // collect at most 2 lines
    }
    if (nameParts.length > 0) return nameParts.join('\n');
  }

  // Strategy 2: generic label search (for colon-format UDDs)
  const candidates = ['name', 'author', 'prepared by', 'created by', 'document owner'];
  return findValueByLabels(lines, candidates, looksLikeName);
}

/**
 * Find the line index of the "SAP ECC" platform line near the start of the UDD.
 * Returns -1 if not found within the first 80 lines.
 */
function findSapEccLine(lines) {
  for (let i = 0; i < Math.min(80, lines.length); i++) {
    const t = lines[i].trim().toLowerCase();
    if (t.startsWith('sap ecc') || t === 'sap ecc/6.0' || t === 'sap ecc 6.0') return i;
  }
  return -1;
}

/** Returns true if the string is a role/function label, not a person name */
function isRoleLabel(str) {
  const norm = normalizeLabel(str);
  const roleLabels = ['developer', 'functional', 'coordinator', 'reviewer', 'guardian',
                      'module owner', 'author', 'co-author', 'function', 'manager'];
  return roleLabels.some(r => norm === r || norm.startsWith(r + ' '));
}

/**
 * Extract CRR Document Title.
 * The UDD typically contains a UDD-... title. We convert it to CRR-... format.
 * Also looks for explicit CRR-... codes anywhere in the document.
 */
function extractCRRTitle(lines) {
  // First priority: find an explicit CRR-... code anywhere
  const crrPattern = /\bCRR[-_][A-Z0-9][-A-Z0-9_]{2,}\b/i;
  for (const line of lines) {
    const match = line.match(crrPattern);
    if (match) {
      return match[0].trim().toUpperCase();
    }
  }

  // Second priority: find a UDD-... title and convert to CRR-...
  const uddPattern = /\bUDD[-_]([A-Z0-9][-A-Z0-9_]{2,})\b/i;
  for (const line of lines) {
    const match = line.match(uddPattern);
    if (match) {
      // Replace UDD prefix with CRR
      return ('CRR-' + match[1]).toUpperCase();
    }
  }

  // Third: look for document title label containing a document code
  const candidates = ['document title', 'crr document title', 'crr title', 'document name', 'doc title', 'udd title'];
  const val = findValueByLabels(lines, candidates);
  if (val) {
    // If it has UDD prefix, convert it
    const uddMatch = val.match(/\bUDD[-_]([A-Z0-9][-A-Z0-9_]{2,})\b/i);
    if (uddMatch) return ('CRR-' + uddMatch[1]).toUpperCase();
    // If it's already CRR
    if (crrPattern.test(val)) return val.match(crrPattern)[0].toUpperCase();
    // Return as-is only if it looks like a document code
    if (/CRR/i.test(val)) return val;
  }

  return null;
}

/**
 * Extract UDD Creation Date — must look like a date value.
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

  function looksLikeDate(str) {
    // Accept formats: 25-May-2022, 17-AUG-2022, 15.08.2026, 2022-05-25, 01/01/2025, etc.
    return /\d/.test(str) && (
      /\d{1,2}[\-\.\/]\w+[\-\.\/]\d{2,4}/.test(str) ||
      /\d{4}[\-\.\/]\d{1,2}[\-\.\/]\d{1,2}/.test(str) ||
      /\d{1,2}[\-\s]\w{3,}[\-\s]\d{2,4}/i.test(str) ||
      /\w{3,}\s+\d{1,2},?\s+\d{4}/i.test(str)
    );
  }

  return findValueByLabels(lines, candidates, looksLikeDate);
}

/**
 * Extract Development Type — must be a text value (not a number).
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
  return findValueByLabels(lines, candidates, v => !/^\d+$/.test(v.trim()) && /[a-zA-Z]/.test(v));
}

/**
 * Extract Reviewer.
 * Handles both "Reviewer: Jane Doe" format AND 3-column tables:
 *   REVIEWER | Coordinator | Shamik Das (E633074)
 */
function extractReviewer(lines) {
  // Strategy 1: 3-column table pattern: REVIEWER \t function \t name
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    const parts = line.split(/\t+/);
    if (parts.length >= 3) {
      const roleCell = normalizeLabel(parts[0]);
      if (roleCell === 'reviewer') {
        const name = parts[parts.length - 1].trim();
        if (looksLikeName(name)) return name;
      }
    }
    if (parts.length === 2) {
      const roleCell = normalizeLabel(parts[0]);
      if (roleCell === 'reviewer') {
        const name = parts[1].trim();
        if (looksLikeName(name)) return name;
      }
    }
  }

  // Strategy 2: line-by-line mammoth format — find "REVIEWED BY (CO)" in roles section
  // Structure: "REVIEWED BY (CO)" → empty lines → function value → empty → name
  const rolesStart = findRolesSectionStart(lines);
  const searchStart = rolesStart !== -1 ? rolesStart : 0;
  for (let i = searchStart; i < lines.length; i++) {
    const norm = normalizeLabel(lines[i]);
    if (norm.includes('reviewed by') || norm === 'reviewer') {
      // Skip function value lines; find the name (skip "Reviewer", "Coordinator" etc.)
      for (let j = i + 1; j < Math.min(i + 11, lines.length); j++) {
        const candidate = lines[j].trim();
        if (candidate && looksLikeName(candidate) && !isRoleLabel(candidate)) {
          return candidate;
        }
      }
    }
  }

  // Strategy 3: generic label search
  const candidates = ['reviewer', 'reviewed by', 'review by', 'code reviewer', 'technical reviewer'];
  return findValueByLabels(lines, candidates, looksLikeName);
}

/**
 * Extract Developer Function.
 * Handles both "Function: Developer" format AND 3-column tables:
 *   DEVELOPER | Developer | Christian Khouri (E631475)
 * The FUNCTION column (middle) should be extracted, not a number.
 */
function extractDeveloperFunction(lines) {
  // Strategy 1: look for 3-column table row with DEVELOPER label
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    const parts = line.split(/\t+/);
    if (parts.length >= 3) {
      const roleCell = normalizeLabel(parts[0]);
      if (roleCell === 'developer' || roleCell === 'dev') {
        // parts[1] = function (middle column)
        const fn = parts[1].trim();
        if (looksLikeFunctionValue(fn)) return fn;
      }
    }
  }

  // Strategy 2: line-by-line mammoth format — find CO-AUTHOR (DEV) label in roles section.
  // Structure: "CO-AUTHOR (DEV)" → empty lines → function value → empty → name
  const rolesStart = findRolesSectionStart(lines);
  if (rolesStart !== -1) {
    for (let i = rolesStart; i < lines.length; i++) {
      const norm = normalizeLabel(lines[i]);
      if (norm.includes('co-author') && norm.includes('dev')) {
        // First non-empty, non-role-label value is the function
        for (let j = i + 1; j < Math.min(i + 11, lines.length); j++) {
          const candidate = lines[j].trim();
          if (!candidate) continue;
          if (isRoleLabelStrict(candidate) && !looksLikeFunctionValue(candidate)) break;
          if (looksLikeFunctionValue(candidate)) return candidate;
        }
      }
    }
  }

  // Strategy 3: generic label fallback (handles "Developer Function: value" and tab formats)
  const candidates = ['developer function', 'function', 'developer role', 'designation', 'job title', 'position'];
  return findValueByLabels(lines, candidates, looksLikeFunctionValue);
}

/**
 * Extract Developer Name.
 * Handles both label format AND 3-column table:
 *   DEVELOPER | Developer | Christian Khouri (E631475)
 */
function extractDeveloperName(lines) {
  // Strategy 1: 3-column table — DEVELOPER | function | NAME
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    const parts = line.split(/\t+/);
    if (parts.length >= 3) {
      const roleCell = normalizeLabel(parts[0]);
      if (roleCell === 'developer' || roleCell === 'dev') {
        // Last column is the name
        const name = parts[parts.length - 1].trim();
        if (looksLikeName(name)) return name;
      }
    }
    // 2-column: DEVELOPER | name
    if (parts.length === 2) {
      const roleCell = normalizeLabel(parts[0]);
      if (roleCell === 'developer' || roleCell === 'dev') {
        const name = parts[1].trim();
        if (looksLikeName(name)) return name;
      }
    }
  }

  // Strategy 2: line-by-line mammoth format — find CO-AUTHOR (DEV) in roles section
  const rolesStart2 = findRolesSectionStart(lines);
  if (rolesStart2 !== -1) {
    for (let i = rolesStart2; i < lines.length; i++) {
      const norm = normalizeLabel(lines[i]);
      if (norm.includes('co-author') && norm.includes('dev')) {
        // Skip next lines until we find a name (not a role/function label)
        for (let j = i + 1; j < Math.min(i + 11, lines.length); j++) {
          const candidate = lines[j].trim();
          if (candidate && looksLikeName(candidate) && !isRoleLabel(candidate)) {
            return candidate;
          }
        }
      }
    }
  }

  // Strategy 3: generic label fallback (handles "Developer Name: value" and tab formats)
  const candidates = ['developer name', 'developer', 'developed by', 'programmer', 'abap developer'];
  return findValueByLabels(lines, candidates, looksLikeName);
}

/**
 * Find the line index where the "Roles and responsibilities" section starts.
 * Returns -1 if not found.
 */
function findRolesSectionStart(lines) {
  for (let i = 0; i < lines.length; i++) {
    const norm = normalizeLabel(lines[i]);
    if (norm === 'roles and responsibilities' || norm === 'roles and responsibility' ||
        norm === '1.2 roles and responsibilities' || norm === '1.2 roles and responsibility') {
      return i;
    }
  }
  return -1;
}

/**
 * Stricter role label check — exact match only (for loop termination)
 */
function isRoleLabelStrict(str) {
  const norm = normalizeLabel(str);
  const strict = ['reviewer', 'developer', 'module owner', 'guardian', 'author (fc)',
                  'co-author (dev)', 'reviewed by (co)', 'function', 'name'];
  return strict.some(r => norm === r || norm.startsWith(r + ' (') || norm === r.replace(' ', '-'));
}

// ──────────────────────────────────────────────────────────────────────────────
// Main extraction function
// ──────────────────────────────────────────────────────────────────────────────

/**
 * @param {string} rawText - plain text extracted from UDD by mammoth
 * @returns {Object} extracted fields with null for missing ones
 */
function extractFieldsFromUDD(rawText) {
  const lines = rawText.split('\n');

  const name            = extractName(lines);
  const crrTitle        = extractCRRTitle(lines);
  const uddCreationDate = extractUDDCreationDate(lines);
  const developmentType = extractDevelopmentType(lines);
  const reviewer        = extractReviewer(lines);
  const developerFunction = extractDeveloperFunction(lines);
  const developerName   = extractDeveloperName(lines);

  return { name, crrTitle, uddCreationDate, developmentType, reviewer, developerFunction, developerName };
}

/**
 * Validate extraction results.
 * Returns an array of error strings (empty = all good).
 */
function validateExtraction(fields) {
  const errors = [];
  const fieldNames = {
    name:             'Name',
    crrTitle:         'CRR Document Title',
    uddCreationDate:  'UDD Creation Date',
    developmentType:  'Development Type',
    reviewer:         'Reviewer',
    developerFunction:'Developer Function',
    developerName:    'Developer Name',
  };
  for (const [key, label] of Object.entries(fieldNames)) {
    if (!fields[key]) {
      errors.push(`"${label}" could not be identified in the uploaded UDD.`);
    }
  }
  return errors;
}

module.exports = { extractFieldsFromUDD, validateExtraction };
