'use strict';

const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const fse = require('fs-extra');
const mammoth = require('mammoth');
const { v4: uuidv4 } = require('uuid');

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

// ──────────────────────────────────────────────────────────────────────────────
// In-memory session store (keyed by session UUID)
// Sessions hold: uddPath, crrPath, crrOriginalName, extractedFields, generatedPath
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
  if (sess.uddPath) fse.removeSync(sess.uddPath);
  if (sess.crrPath) fse.removeSync(sess.crrPath);
  if (sess.generatedPath) fse.removeSync(sess.generatedPath);
  sessions.delete(id);
}

// ──────────────────────────────────────────────────────────────────────────────
// Route: POST /api/upload
// Accepts both UDD and CRR files, creates a session, extracts fields from UDD.
// ──────────────────────────────────────────────────────────────────────────────
app.post(
  '/api/upload',
  upload.fields([
    { name: 'udd', maxCount: 1 },
    { name: 'crr', maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const uddFile = req.files?.udd?.[0];
      const crrFile = req.files?.crr?.[0];

      if (!uddFile) {
        cleanupUploadedFiles([crrFile]);
        return res.status(400).json({ error: 'Please upload the UDD document.' });
      }
      if (!crrFile) {
        cleanupUploadedFiles([uddFile]);
        return res.status(400).json({ error: 'Please upload the CRR template.' });
      }

      // Extract text from UDD using mammoth (preserves table structure via newlines)
      let uddText;
      try {
        const result = await mammoth.extractRawText({ path: uddFile.path });
        uddText = result.value;
      } catch (e) {
        cleanupUploadedFiles([uddFile, crrFile]);
        return res.status(422).json({ error: 'Unable to read the UDD document. Please ensure it is a valid DOCX file.' });
      }

      // Validate CRR is readable
      try {
        const PizZip = require('pizzip');
        const buf = fs.readFileSync(crrFile.path);
        new PizZip(buf); // will throw if invalid
      } catch (e) {
        cleanupUploadedFiles([uddFile, crrFile]);
        return res.status(422).json({ error: 'Unable to read the CRR template. Please ensure it is a valid DOCX file.' });
      }

      // Extract fields from UDD
      const fields = extractFieldsFromUDD(uddText);
      const errors = validateExtraction(fields);

      // Create session
      const sessionId = uuidv4();
      sessions.set(sessionId, {
        createdAt: Date.now(),
        uddPath: uddFile.path,
        crrPath: crrFile.path,
        crrOriginalName: crrFile.originalname,
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

    // Use confirmed fields (user may have corrected values)
    const finalFields = fields || sess.extractedFields;

    // Validate all required fields are present
    const errors = validateExtraction(finalFields);
    if (errors.length > 0) {
      return res.status(422).json({ errors });
    }

    // Read CRR buffer
    let crrBuffer;
    try {
      crrBuffer = fs.readFileSync(sess.crrPath);
    } catch (e) {
      return res.status(500).json({ error: 'CRR template file is no longer available. Please re-upload.' });
    }

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

    // Build output filename
    const crrBase = path.basename(sess.crrOriginalName, '.docx');
    // Use CRR title from fields if available, otherwise use CRR filename
    const titleBase = finalFields.crrTitle || crrBase;
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
