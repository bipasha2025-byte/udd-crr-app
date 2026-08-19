'use strict';
const PizZip = require('pizzip');
const fs = require('fs');

const uploads = fs.readdirSync('./uploads')
  .map(f => ({ f, t: fs.statSync('./uploads/' + f).mtime }))
  .sort((a, b) => b.t - a.t);
const zip = new PizZip(fs.readFileSync('./uploads/' + uploads[0].f));
const doc = zip.file('word/document.xml').asText();
const rows = doc.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || [];

function extractText(xml) {
  return (xml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || []).map(t => t.replace(/<[^>]+>/g, '')).join('');
}
function splitCells(rowXml) {
  const re = /(<w:tc[ >][\s\S]*?<\/w:tc>|<w:tc>[\s\S]*?<\/w:tc>)/g;
  let m, cells = [];
  while ((m = re.exec(rowXml)) !== null) cells.push({ xml: m[1], index: m.index, length: m[1].length });
  return cells;
}
function replaceCell(rowXml, cellObj, newXml) {
  return rowXml.substring(0, cellObj.index) + newXml + rowXml.substring(cellObj.index + cellObj.length);
}
function injectValue(cellXml, value) {
  const enc = value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  let replaced = false;
  const result = cellXml.replace(/<w:t([^>]*)>([^<]*)<\/w:t>/g, (full, attrs, inner) => {
    if (!replaced) { replaced = true; return `<w:t${attrs} xml:space="preserve">${enc}</w:t>`; }
    return `<w:t${attrs}></w:t>`;
  });
  if (!replaced) {
    const lastP = result.lastIndexOf('</w:p>');
    if (lastP >= 0) return result.substring(0, lastP) + `<w:r><w:t xml:space="preserve">${enc}</w:t></w:r>` + result.substring(lastP);
    return result.replace('</w:tc>', `<w:p><w:r><w:t xml:space="preserve">${enc}</w:t></w:r></w:p></w:tc>`);
  }
  return result;
}

// Simulate processTableRow on row 7
const row7 = rows[7];
console.log('Row 7 cells:');
const cells = splitCells(row7);
cells.forEach((c, i) => console.log(`  Cell ${i} index:${c.index} text:${JSON.stringify(extractText(c.xml))}`));

// Step 1: inject into cells[1]
let newRow = replaceCell(row7, cells[1], injectValue(cells[1].xml, 'Developer'));
console.log('\nAfter Step1 inject (fn), row length:', newRow.length, '(was', row7.length + ')');

// Step 2: reparse and inject into updatedCells[2]
const cells2 = splitCells(newRow);
console.log('Reparsed cells:');
cells2.forEach((c, i) => console.log(`  Cell ${i} index:${c.index} text:${JSON.stringify(extractText(c.xml))}`));

newRow = replaceCell(newRow, cells2[2], injectValue(cells2[2].xml, 'Christian Khouri (E631475)'));
const cells3 = splitCells(newRow);
console.log('\nFinal row cells:', JSON.stringify(cells3.map(c => extractText(c.xml))));
