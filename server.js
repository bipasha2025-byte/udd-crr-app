'use strict';

const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const fse = require('fs-extra');
const mammoth = require('mammoth');
const { v4: uuidv4 } = require('uuid');

// ──────────────────────────────────────────────────────────────────────────────
// Internal CRR master template — stored in templates/ at deploy time.
// Loaded once at startup; never modified at runtime.
// ──────────────────────────────────────────────────────────────────────────────
const CRR_TEMPLATE_PATH = path.join(__dirname, 'templates', 'CRR_Master_Template.docx');
if (!fs.existsSync(CRR_TEMPLATE_PATH)) {
  console.error('FATAL: CRR master template not found at', CRR_TEMPLATE_PATH);
  process.exit(1);
}
const CRR_MASTER_BUFFER = fs.readFileSync(CRR_TEMPLATE_PATH);
console.log(`CRR master template loaded: ${CRR_TEMPLATE_PATH} (${CRR_MASTER_BUFFER.length} bytes)`);

/**
 * Returns today's date formatted as DD-Mon-YYYY (e.g. "15-Jun-2025").
 * This is always later than any UDD creation date from the past.
 */
function todayFormatted() {
  const now = new Date();
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const dd = String(now.getDate()).padStart(2, '0');
  const mon = months[now.getMonth()];
  const yyyy = now.getFullYear();
  return `${dd}-${mon}-${yyyy}`;
}

const { extractFieldsFromUDD, validateExtraction } = require('./extractor');
const { populateCRR } = require('./populator');

const app = express();
const PORT = process.env.PORT || 3000;

// ──────────────────────────────────────────────────────────────────────────────
// Directories — use /tmp on cloud (Railway), local dirs in development
// ──────────────────────────────────────────────────────────────────────────────
const IS_CLOUD = process.env.RAILWAY_ENVIRONMENT || process.env.RENDER || process.env.NODE_ENV === 'production';
const BASE_TMP = IS_CLOUD ? '/tmp' : __dirname;
const UPLOADS_DIR = path.join(BASE_TMP, 'uploads');
const GENERATED_DIR = path.join(BASE_TMP, 'generated');
fse.ensureDirSync(UPLOADS_DIR);
fse.ensureDirSync(GENERATED_DIR);

// ──────────────────────────────────────────────────────────────────────────────
// Middleware
// ──────────────────────────────────────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Multer — store files with unique names, allow only .docx
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOADS_DIR),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    cb(null, `${uuidv4()}${ext}`);
  },
});

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext !== '.docx') {
      return cb(new Error('Please upload a valid DOCX document.'));
    }
    cb(null, true);
  },
  limits: { fileSize: 50 * 1024 * 1024 }, // 50 MB
});

// UDD-only upload handler — only the 'udd' field is accepted now
const uploadUDD = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext !== '.docx') {
      return cb(new Error('Please upload a valid DOCX document.'));
    }
    cb(null, true);
  },
  limits: { fileSize: 50 * 1024 * 1024 },
});

// ──────────────────────────────────────────────────────────────────────────────
// In-memory session store (keyed by session UUID)
// Sessions hold: uddPath, extractedFields, generatedPath
// (crrPath is no longer stored — the internal template is always used)
// ──────────────────────────────────────────────────────────────────────────────
const sessions = new Map();

function getSession(id) {
  return sessions.get(id);
}

function setSession(id, data) {
  sessions.set(id, { ...sessions.get(id), ...data });
}

// Cleanup old sessions (> 1 hour)
setInterval(() => {
  const now = Date.now();
  for (const [id, sess] of sessions.entries()) {
    if (now - sess.createdAt > 60 * 60 * 1000) {
      cleanupSession(id);
    }
  }
}, 10 * 60 * 1000);

function cleanupSession(id) {
  const sess = sessions.get(id);
  if (!sess) return;
  if (sess.uddPath)      fse.removeSync(sess.uddPath);
  if (sess.generatedPath) fse.removeSync(sess.generatedPath);
  // No crrPath to clean — the internal template is never written per-session
  sessions.delete(id);
}

// ──────────────────────────────────────────────────────────────────────────────
// Route: POST /api/upload
// Accepts only the UDD file. The CRR master template is loaded internally.
// ──────────────────────────────────────────────────────────────────────────────
app.post(
  '/api/upload',
  uploadUDD.single('udd'),
  async (req, res) => {
    try {
      const uddFile = req.file;

      if (!uddFile) {
        return res.status(400).json({ error: 'Please upload the UDD document.' });
      }

      // Extract text from UDD using mammoth
      let uddText;
      try {
        const result = await mammoth.extractRawText({ path: uddFile.path });
        uddText = result.value;
      } catch (e) {
        fse.removeSync(uddFile.path);
        return res.status(422).json({ error: 'Unable to read the UDD document. Please ensure it is a valid DOCX file.' });
      }

      // Read raw buffer for XML-direct extraction (App Components table)
      const uddBuffer = fs.readFileSync(uddFile.path);

      // Extract fields from UDD — pass buffer so App Components uses XML path
      const fields = extractFieldsFromUDD(uddText, uddBuffer);
      const errors = validateExtraction(fields);

      // Derive relatedUDDName from the original UDD filename — strip only the .docx extension,
      // preserve the rest exactly (including document numbers, hyphens, suffixes like "-04").
      // e.g. "UDD-SAPECC-QM-1609-04.docx" → "UDD-SAPECC-QM-1609-04"
      const uddBaseName = path.basename(uddFile.originalname, path.extname(uddFile.originalname));
      fields.relatedUDDName = uddBaseName;

      // Create session — no crrPath needed; the internal master template is used at generate time
      const sessionId = uuidv4();
      sessions.set(sessionId, {
        createdAt: Date.now(),
        uddPath: uddFile.path,
        extractedFields: fields,
        generatedPath: null,
      });

      if (errors.length > 0) {
        return res.status(200).json({
          sessionId,
          fields,
          warnings: errors,
          success: false,
        });
      }

      return res.status(200).json({
        sessionId,
        fields,
        warnings: [],
        success: true,
      });
    } catch (err) {
      console.error('Upload error:', err);
      return res.status(500).json({ error: 'An unexpected error occurred during upload.' });
    }
  }
);

function cleanupUploadedFiles(files) {
  for (const f of files) {
    if (f && f.path) fse.removeSync(f.path);
  }
}

// ──────────────────────────────────────────────────────────────────────────────
// Route: POST /api/generate
// Receives confirmed fields + sessionId, generates the CRR.
// ──────────────────────────────────────────────────────────────────────────────
app.post('/api/generate', express.json(), async (req, res) => {
  try {
    const { sessionId, fields } = req.body;

    if (!sessionId || !sessions.has(sessionId)) {
      return res.status(400).json({ error: 'Invalid or expired session. Please re-upload the documents.' });
    }

    const sess = getSession(sessionId);

    // Merge: start from session's extracted fields (contains server-injected values like relatedUDDName),
    // then overlay any user-confirmed edits on top.
    // reviewType comes directly from the request body (user selection from UI).
    const finalFields = Object.assign({}, sess.extractedFields, fields || {});
    if (req.body.reviewType) finalFields.reviewType = req.body.reviewType;

    // Validate all required fields are present
    const errors = validateExtraction(finalFields);
    if (errors.length > 0) {
      return res.status(422).json({ errors });
    }

    // Use the internal CRR master template buffer (loaded once at startup)
    const crrBuffer = CRR_MASTER_BUFFER;

    // Inject CRR Creation Date = today's date (always later than the UDD creation date)
    finalFields.crrCreationDate = todayFormatted();

    // Generate populated CRR
    let outputBuffer;
    try {
      outputBuffer = populateCRR(crrBuffer, finalFields);
    } catch (e) {
      console.error('Population error:', e);
      return res.status(500).json({ error: 'Failed to populate the CRR template: ' + e.message });
    }

    // Build output filename — use the extracted CRR title if available,
    // otherwise fall back to the internal template name
    const titleBase = finalFields.crrTitle || 'CRR';
    const safeTitle = titleBase.replace(/[^a-zA-Z0-9\-_]/g, '_');
    const outputFilename = `${safeTitle}_Filled.docx`;
    const outputPath = path.join(GENERATED_DIR, `${sessionId}_${outputFilename}`);

    fs.writeFileSync(outputPath, outputBuffer);
    setSession(sessionId, { generatedPath: outputPath, outputFilename });

    return res.status(200).json({ sessionId, outputFilename, success: true });
  } catch (err) {
    console.error('Generate error:', err);
    return res.status(500).json({ error: 'An unexpected error occurred during generation.' });
  }
});

// ──────────────────────────────────────────────────────────────────────────────
// Route: GET /api/download/:sessionId
// ──────────────────────────────────────────────────────────────────────────────
app.get('/api/download/:sessionId', (req, res) => {
  const sess = sessions.get(req.params.sessionId);
  if (!sess || !sess.generatedPath || !fs.existsSync(sess.generatedPath)) {
    return res.status(404).json({ error: 'File not found or session expired.' });
  }

  res.setHeader('Content-Disposition', `attachment; filename="${sess.outputFilename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.sendFile(sess.generatedPath, (err) => {
    if (!err) {
      // Schedule cleanup after download
      setTimeout(() => cleanupSession(req.params.sessionId), 5 * 60 * 1000);
    }
  });
});

// ──────────────────────────────────────────────────────────────────────────────
// Route: GET /health  — used by Railway / Render to confirm app is alive
// ──────────────────────────────────────────────────────────────────────────────
app.get('/health', (req, res) => res.json({ status: 'ok' }));

// ──────────────────────────────────────────────────────────────────────────────
// Route: DELETE /api/session/:sessionId  (optional cleanup)
// ──────────────────────────────────────────────────────────────────────────────
app.delete('/api/session/:sessionId', (req, res) => {
  cleanupSession(req.params.sessionId);
  res.json({ ok: true });
});

// ──────────────────────────────────────────────────────────────────────────────
// Error handler for multer file-type errors
// ──────────────────────────────────────────────────────────────────────────────
app.use((err, req, res, next) => {
  if (err && err.message) {
    return res.status(400).json({ error: err.message });
  }
  next(err);
});

// ──────────────────────────────────────────────────────────────────────────────
// Start server
// ──────────────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`UDD-to-CRR Automation running at http://localhost:${PORT}`);
});

module.exports = app;
