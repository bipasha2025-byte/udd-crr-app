'use strict';

/**
 * debug_name.js — trace exactly why extractName() fails on the real UDD
 * Usage: node debug_name.js
 */

const mammoth = require('mammoth');
const fs = require('fs');
const path = require('path');
const { extractFieldsFromUDD } = require('./extractor');

const uploadDir = './uploads';
const files = fs.readdirSync(uploadDir)
  .map(f => ({ f, t: fs.statSync(path.join(uploadDir, f)).mtime }))
  .sort((a, b) => b.t - a.t);

if (files.length < 2) {
  console.error('Need at least 2 files in ./uploads (UDD and CRR).');
  process.exit(1);
}

// The UDD is likely the second-most-recent file (or we test both)
async function run() {
  for (const { f } of files) {
    const fpath = path.join(uploadDir, f);
    console.log('\n========================================');
    console.log('FILE:', f);
    try {
      const result = await mammoth.extractRawText({ path: fpath });
      const lines = result.value.split('\n');

      // Show lines 160–210 (area of interest)
      console.log('\n--- Lines 155-220 ---');
      for (let i = 155; i < Math.min(225, lines.length); i++) {
        const repr = JSON.stringify(lines[i]);
        console.log(`  [${i}] ${repr}`);
      }

      // Show first 20 lines too
      console.log('\n--- Lines 0-30 ---');
      for (let i = 0; i < Math.min(30, lines.length); i++) {
        console.log(`  [${i}] ${JSON.stringify(lines[i])}`);
      }

      // Show lines containing 'co-author', 'developer', 'reviewer', 'author'
      console.log('\n--- Lines containing key role terms ---');
      for (let i = 0; i < lines.length; i++) {
        const low = lines[i].toLowerCase();
        if (low.includes('co-author') || low.includes('author') || low.includes('reviewer') || low.includes('developer')) {
          console.log(`  [${i}] ${JSON.stringify(lines[i])}`);
        }
      }

      // Run extraction
      console.log('\n--- Extraction result ---');
      const fields = extractFieldsFromUDD(result.value);
      console.log(JSON.stringify(fields, null, 2));

    } catch (e) {
      console.log('  Error:', e.message);
    }
  }
}

run().catch(console.error);
