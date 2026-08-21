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
// Section 4 field extractors (new)
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Find the line index of "System Components" section (UDD 7.2).
 * Searches within a reasonable range, not the whole doc.
 */
function findSystemComponentsStart(lines) {
  for (let i = 0; i < lines.length; i++) {
    const norm = normalizeLabel(lines[i]);
    if (norm === 'system components' || norm === '7.2 system components' ||
        norm === '7.2. system components') return i;
  }
  return -1;
}

/**
 * Extract a labelled value from the System Components section.
 * Structure (mammoth line-by-line):
 *   "R/3 Version:"   ← label line (may include colon)
 *   "SAP ECC 6.0"    ← value line
 */
function extractFromSystemComponents(lines, labelPattern, allowNA) {
  const sysStart = findSystemComponentsStart(lines);
  if (sysStart === -1) return null;
  // Search within 60 lines of the section heading
  for (let i = sysStart; i < Math.min(sysStart + 60, lines.length); i++) {
    const norm = normalizeLabel(lines[i]);
    if (labelPattern.test(norm)) {
      // Value may be on same line after colon, or on next non-empty line
      const colonIdx = lines[i].indexOf(':');
      if (colonIdx !== -1) {
        const inline = lines[i].substring(colonIdx + 1).trim();
        // allowNA = treat "N/A" as valid (e.g. Legacy System: N/A)
        if (inline && (allowNA || !isPlaceholder(inline))) return inline;
      }
      // Look at next 1-4 lines
      for (let j = i + 1; j < Math.min(i + 5, lines.length); j++) {
        const t = lines[j].trim();
        if (t && (allowNA || !isPlaceholder(t)) && !looksLikeLabel(t)) return t;
      }
    }
  }
  return null;
}

/** Extract R/3 Version from UDD section 7.2 */
function extractR3Version(lines) {
  return extractFromSystemComponents(lines, /r[\s\/]?3\s*version/i) ||
         findValueByLabels(lines, ['r/3 version', 'r3 version', 'sap version'], null);
}

/** Extract Source System from UDD section 7.2 */
function extractSourceSystem(lines) {
  return extractFromSystemComponents(lines, /source\s*system/i) ||
         findValueByLabels(lines, ['source system'], null);
}

/** Extract Legacy System from UDD section 7.2 */
function extractLegacySystem(lines) {
  // "N/A" IS a valid legacy system value — don't treat as placeholder here
  return extractFromSystemComponents(lines, /legacy\s*system/i, true) ||
         findValueByLabels(lines, ['legacy system'], null);
}

/**
 * Find the ACTUAL Appendix 1 / Revision Log section start (not the TOC reference).
 * We require that "DOCUMENT VERSION" or "DATE OF THE CHANGE" header appears
 * within 30 lines after the match — this distinguishes the real section from TOC.
 */
function findRevisionLogStart(lines) {
  for (let i = 0; i < lines.length; i++) {
    const t = lines[i].trim();
    if (/appendix\s*1.*revision\s*log/i.test(t) || /^revision\s*log$/i.test(t)) {
      // Verify this is the actual section (has header within 30 lines)
      for (let j = i + 1; j < Math.min(i + 30, lines.length); j++) {
        if (/document version/i.test(lines[j]) || /date of the change/i.test(lines[j])) return i;
      }
    }
  }
  return -1;
}

/**
 * Parse the revision log into entries.
 * Each entry is keyed by columns: version, date, reasons, transport, etc.
 * Returns array of { version, date, reasons, projectName, crqNumbers[] }
 */
function parseRevisionLog(lines) {
  const revStart = findRevisionLogStart(lines);
  if (revStart === -1) return [];

  // Find header row (DOCUMENT VERSION / DATE OF THE CHANGE / REASONS OF THE CHANGE)
  let headerIdx = -1;
  for (let i = revStart; i < Math.min(revStart + 20, lines.length); i++) {
    if (/document version/i.test(lines[i])) { headerIdx = i; break; }
  }
  if (headerIdx === -1) return [];

  // Collect all non-empty lines after header — the table is rendered cell-by-cell
  const dataLines = [];
  for (let i = headerIdx + 1; i < lines.length; i++) {
    const t = lines[i].trim();
    if (t) dataLines.push(t);
  }

  // The revision log table columns (mammoth line order): version, date, reasons, transport, (dates/authors)
  // Entries are separated by version numbers like "01", "02", "03", "04"
  const entries = [];
  let current = null;

  for (let i = 0; i < dataLines.length; i++) {
    const t = dataLines[i];
    // Detect version number (1-2 digit number alone, possibly like "01", "02", "03")
    if (/^\d{1,2}$/.test(t) && i + 1 < dataLines.length) {
      // Save previous entry
      if (current) entries.push(current);
      current = { version: t, date: null, reasons: null, projectName: null, crqNumbers: [] };
      continue;
    }
    if (!current) continue;

    // Date — looks like a date and no date recorded yet
    if (!current.date && looksLikeDateStr(t)) {
      current.date = t;
      continue;
    }

    // CRQ numbers — extract any CRQ patterns in this line
    const crqMatches = t.match(/CRQ\s*\d{6,}/gi) || t.match(/CRQ\d+/gi) || [];
    for (const crq of crqMatches) {
      const norm = crq.replace(/\s+/g, '').toUpperCase();
      if (!current.crqNumbers.includes(norm)) current.crqNumbers.push(norm);
    }

    // Project name — the FIRST standalone heading line in the REASONS OF THE CHANGE cell,
    // appearing before the detailed change description.
    // Rules:
    //   - Must be the first qualifying line (once set, do not overwrite)
    //   - Must contain at least 3 letters
    //   - Must NOT be a transport request ID (DE1K..., AE1K..., PE1K...)
    //   - Must NOT be a date string
    //   - Must NOT be a version number alone
    //   - Must NOT be an author/update line ("TS updated by", "FS Updated by", etc.)
    //   - Must NOT be a CRQ-only line (a line whose ONLY content is a CRQ number)
    //   - Must NOT be an incident-only line (INC followed by digits only)
    //   - Short names like "UCB_RUN" or "UCB RUN" ARE valid — no minimum length
    if (!current.projectName &&
        !/^DE1K|^AE1K|^PE1K|^[A-Z]{2,3}K\d/i.test(t) &&
        !looksLikeDateStr(t) &&
        !/^\d{1,2}$/.test(t) &&
        !/^(end|begin|start)\s*of\s*crq/i.test(t) &&
        !/^(ts|fs)\s+(updated|update)\s+by/i.test(t) &&
        !/^(formal\s+adjustments|kernel\s+release|TS\s+updated|FS\s+updated)/i.test(t) &&
        !/^INC\d/i.test(t) &&
        /[a-zA-Z]{3}/.test(t)) {
      // A line that is ONLY a CRQ number should not become the project name
      const isCRQOnly = /^CRQ[\s\d]+$/i.test(t.trim());
      if (!isCRQOnly) {
        current.projectName = t;
      }
    }
  }
  if (current) entries.push(current);
  return entries;
}

function looksLikeDateStr(str) {
  return /\d/.test(str) && (
    /\d{1,2}[\-\.\/]\w+[\-\.\/]\d{2,4}/.test(str) ||
    /\d{4}[\-\.\/]\d{1,2}[\-\.\/]\d{1,2}/.test(str) ||
    /\d{1,2}[\-\s]\w{3,}[\-\s]\d{2,4}/i.test(str)
  );
}

/**
 * Extract CRQ Number — from the latest (highest version) revision log entry.
 * If multiple CRQs found in latest entry, return the most prominent one.
 */
function extractCRQNumber(lines) {
  const entries = parseRevisionLog(lines);
  if (entries.length === 0) return null;
  // Use the last entry (latest revision)
  const latest = entries[entries.length - 1];
  if (latest.crqNumbers.length === 0) return null;
  // Return the first (most prominent) CRQ from the latest entry
  return latest.crqNumbers[0];
}

/**
 * Extract Project Name — from the latest revision log entry.
 */
function extractProjectName(lines) {
  const entries = parseRevisionLog(lines);
  if (entries.length === 0) return null;
  const latest = entries[entries.length - 1];
  return latest.projectName || null;
}

// ──────────────────────────────────────────────────────────────────────────────
// App Components & Objects extractor (UDD 7.2 → CRR "Copied objects" table)
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Known object-type category labels in UDD 7.2 App Components table.
 * Used to detect where a new category block starts.
 */
const APP_COMPONENT_CATEGORIES = new Set([
  'module / package', 'report', 'function group', 'structure', 'basic type',
  'function module', 'include', 'table', 'view', 'class', 'interface',
  'program', 'subroutine pool', 'type group', 'message class', 'domain',
  'data element', 'lock object', 'search help', 'number range',
  'enhancement spot', 'badi definition', 'badi implementation',
  'idoc type', 'message type', 'logical message', 'extension type',
]);

function isAppComponentCategory(line) {
  return APP_COMPONENT_CATEGORIES.has(line.trim().toLowerCase());
}

/**
 * Returns true if a line looks like a SAP object name:
 * starts with Z, Y, or / and contains only uppercase letters, digits, underscores.
 */
function looksLikeSAPObject(line) {
  const t = line.trim();
  if (t.length < 2) return false;
  // SAP custom objects start with Z or Y; some start with /
  return /^[ZY\/][A-Z0-9_\/]{1,}$/i.test(t) && !/^Y$/.test(t);
}

/**
 * Extract App Components & Objects from UDD section 7.2.
 * Returns an array of { name, objectType, comment } objects.
 * - name:       object name (one entry per object, never combined)
 * - objectType: category label from leftmost column of that row
 * - codeVersion: always null (not in UDD 7.2)
 * - comment:    CRQ number from Upgrade Implications column for that specific object,
 *               or null if that object's Upgrade Implications cell is blank
 */
function extractAppComponents(lines) {
  const sysStart = findSystemComponentsStart(lines);
  if (sysStart === -1) return [];

  // Find "App Components & Objects" header within 80 lines of system components
  let appStart = -1;
  for (let i = sysStart; i < Math.min(sysStart + 80, lines.length); i++) {
    if (/app components.*objects/i.test(lines[i])) { appStart = i; break; }
  }
  if (appStart === -1) return [];

  // Find the end of the section (next major section heading)
  let appEnd = lines.length;
  for (let i = appStart + 1; i < lines.length; i++) {
    const t = lines[i].trim().toLowerCase();
    if (/^data description$/i.test(t) || /^7\.\d+/.test(t) || /^8\.\s/i.test(t)) {
      appEnd = i; break;
    }
  }

  // Skip column header lines
  const SKIP_HEADERS = new Set(['name', 'existing', 'new', 'upgrade implications',
                                 'app components & objects', 'app components and objects']);

  // Collect non-empty lines BUT keep track of which raw lines they came from.
  // We need raw-line indices so we can count blank slots in the Upgrade Implications column.
  const nonEmptyLines = [];   // { text, rawIdx }
  for (let i = appStart + 1; i < appEnd; i++) {
    const t = lines[i].trim();
    if (!t) continue;
    if (SKIP_HEADERS.has(t.toLowerCase())) continue;
    nonEmptyLines.push({ text: t, rawIdx: i });
  }

  // For the Upgrade Implications column we need to count blanks between entries.
  // After finding names, Existing, New — we compute the raw-line span for Upgrade
  // by knowing that each column cell has EXACTLY one slot per name row.
  //
  // Strategy: for each category block, use raw line positions to count N blank-or-CRQ
  // slots in the Upgrade Implications column by scanning WITHIN the raw lines range.

  const results = [];
  let ni = 0;  // index into nonEmptyLines

  while (ni < nonEmptyLines.length) {
    const { text: t, rawIdx } = nonEmptyLines[ni];

    if (!isAppComponentCategory(t)) { ni++; continue; }

    const objectType = t;
    const categoryRawIdx = rawIdx;
    ni++;

    // Collect consecutive SAP object names
    const names = [];
    const nameRawIdxs = [];
    while (ni < nonEmptyLines.length && looksLikeSAPObject(nonEmptyLines[ni].text)) {
      names.push(nonEmptyLines[ni].text);
      nameRawIdxs.push(nonEmptyLines[ni].rawIdx);
      ni++;
    }
    if (names.length === 0) continue;

    // Skip Existing column (N "Y"/"N" entries)
    let skipped = 0;
    while (ni < nonEmptyLines.length && skipped < names.length &&
           /^[YNy]$/.test(nonEmptyLines[ni].text)) {
      ni++; skipped++;
    }

    // Skip New column (up to N "Y"/"N" entries)
    let newSkipped = 0;
    while (ni < nonEmptyLines.length && newSkipped < names.length &&
           /^[YNyn]$/.test(nonEmptyLines[ni].text)) {
      ni++; newSkipped++;
    }

    // Collect Upgrade Implications.
    // The Upgrade column has exactly N cells (one per name).
    // Mammoth renders each cell as one line (CRQ value or blank), with one separator
    // blank line between cells. Multiple consecutive blanks at the start are inter-column
    // spacing — skip them to reach the first Upgrade cell.
    //
    // Strategy: collect the raw non-empty lines that belong to the Upgrade column.
    // Those are lines that are either a CRQ code or ENTIRELY blank (cell is empty),
    // where we distinguish cell-blanks by their position relative to the column sequence.
    //
    // Simpler: collect the N values by scanning forward through nonEmptyLines —
    // but use "blank raw lines between non-empties" to detect blank cells.
    // Each upgrade CELL is separated from the previous by at least one blank raw line.
    // If two consecutive non-empty lines appear with NO blank raw line between them,
    // they are in different columns, not both Upgrade.

    const upgrades = [];
    let rawJ = ni < nonEmptyLines.length ? nonEmptyLines[ni].rawIdx : appEnd;

    // Skip any leading blank lines (inter-column gap between New and Upgrade columns)
    while (rawJ < appEnd && !lines[rawJ].trim()) rawJ++;

    while (upgrades.length < names.length && rawJ < appEnd) {
      const raw = lines[rawJ].trim();
      if (!raw) {
        // A blank line here means a blank Upgrade cell for the current name
        upgrades.push(null);
        rawJ++;
        // Skip exactly ONE separator blank line after a blank cell.
        // Do NOT consume multiple blanks — the next blank might be another blank cell.
        if (rawJ < appEnd && !lines[rawJ].trim()) rawJ++;
        continue;
      }
      // Non-empty line
      if (isAppComponentCategory(raw) || SKIP_HEADERS.has(raw.toLowerCase())) break;
      upgrades.push(/CRQ/i.test(raw) ? raw : null);
      rawJ++;
      // Skip exactly ONE separator blank line after a non-empty value.
      // Do NOT consume multiple blanks — the next blank might be a blank cell.
      if (rawJ < appEnd && !lines[rawJ].trim()) rawJ++;
    }

    // Advance ni past any lines we consumed
    while (ni < nonEmptyLines.length && nonEmptyLines[ni].rawIdx < rawJ) ni++;

    // Emit one row per name
    for (let k = 0; k < names.length; k++) {
      results.push({
        name:        names[k],
        objectType:  objectType,
        codeVersion: null,     // not in UDD 7.2
        comment:     upgrades[k] || null,
      });
    }
  }

  return results;
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

  const name              = extractName(lines);
  const crrTitle          = extractCRRTitle(lines);
  const uddCreationDate   = extractUDDCreationDate(lines);
  const developmentType   = extractDevelopmentType(lines);
  const reviewer          = extractReviewer(lines);
  const developerFunction = extractDeveloperFunction(lines);
  const developerName     = extractDeveloperName(lines);

  // Section 4 — REPOSITORY OBJECTS
  const r3Version     = extractR3Version(lines);
  const sourceSystem  = extractSourceSystem(lines);
  const legacySystem  = extractLegacySystem(lines);
  const crqNumber     = extractCRQNumber(lines);
  const projectName   = extractProjectName(lines);
  // relatedUDDName is injected by server.js from the uploaded filename
  // sopConventions and devLanguage are fixed values — populated in populator.js

  const appComponents = extractAppComponents(lines);

  return {
    name, crrTitle, uddCreationDate, developmentType,
    reviewer, developerFunction, developerName,
    r3Version, sourceSystem, legacySystem,
    crqNumber, projectName, appComponents,
  };
}

/**
 * Validate extraction results.
 * Returns an array of error strings (empty = all good).
 * Section 4 fields (r3Version, sourceSystem, legacySystem, crqNumber, projectName)
 * are optional — missing ones are flagged as warnings, not hard errors.
 */
function validateExtraction(fields) {
  const errors = [];
  // Required fields (hard errors)
  const required = {
    name:             'Name',
    crrTitle:         'CRR Document Title',
    uddCreationDate:  'UDD Creation Date',
    developmentType:  'Development Type',
    reviewer:         'Reviewer',
    developerFunction:'Developer Function',
    developerName:    'Developer Name',
  };
  for (const [key, label] of Object.entries(required)) {
    if (!fields[key]) {
      errors.push(`"${label}" could not be identified in the uploaded UDD.`);
    }
  }
  return errors;
}

module.exports = { extractFieldsFromUDD, validateExtraction };
