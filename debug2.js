'use strict';

const mammoth = require('mammoth');
const fs = require('fs');
const path = require('path');

const uploadDir = './uploads';
const files = fs.readdirSync(uploadDir)
  .map(f => ({ f, t: fs.statSync(path.join(uploadDir, f)).mtime }))
  .sort((a, b) => b.t - a.t);

// Find the UDD (has CO-AUTHOR line)
async function run() {
  for (const { f } of files) {
    const fpath = path.join(uploadDir, f);
    const result = await mammoth.extractRawText({ path: fpath });
    const lines = result.value.split('\n');

    // Check if this is the UDD
    if (!lines.some(l => l.includes('CO-AUTHOR (DEV)'))) continue;

    console.log('UDD FILE:', f);

    // Show lines around UDD title
    console.log('\n--- Lines 30-80 (title area) ---');
    for (let i = 30; i < Math.min(80, lines.length); i++) {
      if (lines[i].trim()) console.log(`  [${i}] ${JSON.stringify(lines[i])}`);
    }

    // Show lines around CRR/UDD title references
    console.log('\n--- Lines containing UDD- or CRR- ---');
    for (let i = 0; i < lines.length; i++) {
      if (/UDD[-_]|CRR[-_]/i.test(lines[i])) {
        console.log(`  [${i}] ${JSON.stringify(lines[i])}`);
      }
    }

    // Show lines around "function" labels (for developerFunction bug)
    console.log('\n--- Lines containing "function" (first 20 hits) ---');
    let hits = 0;
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].toLowerCase().includes('function') && hits < 20) {
        console.log(`  [${i}] ${JSON.stringify(lines[i])}`);
        hits++;
      }
    }

    // Show lines 100-165 (section 1.1 area)
    console.log('\n--- Lines 100-165 ---');
    for (let i = 100; i < Math.min(165, lines.length); i++) {
      if (lines[i].trim()) console.log(`  [${i}] ${JSON.stringify(lines[i])}`);
    }

    break;
  }
}

run().catch(console.error);
