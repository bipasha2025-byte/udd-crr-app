'use strict';
const { extractFieldsFromUDD } = require('./extractor');

// Simulate a UDD with the UCB RUN example from the spec
const uddText = [
  '',
  'unit detailed design',
  '',
  'SAP ECC/6.0',
  '',
  'GLIMS INTERFACE',
  '',
  'PROCESS COA DATA REPLY',
  '',
  'UDD Creation date',
  '25-May-2022',
  'Development type',
  'Enhancement',
  '',
  'Roles and responsibilities',
  '',
  'FUNCTION',
  '',
  'NAME',
  '',
  '(plus User ID)',
  '',
  'CO-AUTHOR (DEV)  ',
  '',
  '',
  '',
  'Developer',
  '',
  'John Smith (E123456)',
  '',
  'REVIEWED BY (CO)',
  '',
  '',
  '',
  'Reviewer',
  '',
  'Jane Doe (E654321)',
  '',
  'System Components',
  'R/3 Version:',
  'SAP ECC 6.0',
  'Source System:',
  'DE1/500',
  'Legacy System:',
  'N/A',
  '',
  'Appendix',
  'Revision Log',
  'Appendix 1: Revision Log',
  'DOCUMENT VERSION',
  'DATE OF THE CHANGE',
  'REASONS OF THE CHANGE',
  'Transport Request',
  '01',
  '10-Jan-2020',
  'UCB_RUN',
  'INC0000089997 & INC00000088365_Slowness issue',
  'CRQ00000175469 - System performance issue in inbound delivery',
  'DE1K111111',
].join('\n');

const fields = extractFieldsFromUDD(uddText);
console.log('projectName :', JSON.stringify(fields.projectName));
console.log('crqNumber   :', JSON.stringify(fields.crqNumber));

const pOk = fields.projectName === 'UCB_RUN';
const cOk = fields.crqNumber === 'CRQ00000175469';
console.log('projectName match:', pOk ? 'PASS' : 'FAIL (expected UCB_RUN)');
console.log('crqNumber   match:', cOk ? 'PASS' : 'FAIL (expected CRQ00000175469)');

// Also test multi-entry: latest entry should win
const uddText2 = [
  '',
  'unit detailed design',
  '',
  'SAP ECC/6.0',
  '',
  'SOME PROJECT',
  '',
  'UDD Creation date',
  '01-Jan-2021',
  'Development type',
  'Enhancement',
  '',
  'Roles and responsibilities',
  '',
  'FUNCTION',
  '',
  'NAME',
  '',
  '(plus User ID)',
  '',
  'CO-AUTHOR (DEV)  ',
  '',
  '',
  '',
  'Developer',
  '',
  'Alice Brown (E999)',
  '',
  'REVIEWED BY (CO)',
  '',
  '',
  '',
  'Reviewer',
  '',
  'Bob Green (E888)',
  '',
  'System Components',
  'R/3 Version:',
  'SAP ECC 6.0',
  'Source System:',
  'DE1/500',
  'Legacy System:',
  'N/A',
  '',
  'Appendix',
  'Revision Log',
  'Appendix 1: Revision Log',
  'DOCUMENT VERSION',
  'DATE OF THE CHANGE',
  'REASONS OF THE CHANGE',
  'Transport Request',
  '01',
  '01-Jan-2020',
  'First Project',
  'CRQ000111111',
  'DE1KAAAAAA',
  '02',
  '15-Mar-2022',
  'Latest Project Name',
  'CRQ000222222',
  'DE1KBBBBBB',
].join('\n');

const fields2 = extractFieldsFromUDD(uddText2);
console.log('\nMulti-entry test:');
console.log('projectName :', JSON.stringify(fields2.projectName));
console.log('crqNumber   :', JSON.stringify(fields2.crqNumber));
const p2Ok = fields2.projectName === 'Latest Project Name';
const c2Ok = fields2.crqNumber === 'CRQ000222222';
console.log('projectName match:', p2Ok ? 'PASS' : 'FAIL (expected Latest Project Name)');
console.log('crqNumber   match:', c2Ok ? 'PASS' : 'FAIL (expected CRQ000222222)');

process.exit(pOk && cOk && p2Ok && c2Ok ? 0 : 1);
