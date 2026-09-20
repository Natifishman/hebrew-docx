'use strict';

const test = require('node:test');
const assert = require('node:assert');
const { execSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const { HebrewDoc, heading, p, bullet, table, run, ltr, isolate, isRtl } = require('../src');

/** Unzips document.xml out of a .docx buffer using the system `unzip`. */
async function documentXml(doc) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'hbdocx-'));
  const file = path.join(dir, 'out.docx');
  fs.writeFileSync(file, await doc.toBuffer());
  const xml = execSync(`unzip -p ${JSON.stringify(file)} word/document.xml`, { maxBuffer: 1 << 26 }).toString();
  fs.rmSync(dir, { recursive: true, force: true });
  return xml;
}

test('ltr() wraps text in a directional override', () => {
  assert.strictEqual(ltr('1 : 0..1'), '\u202D1 : 0..1\u202C');
});

test('isolate() protects Latin and numeric segments only', () => {
  const out = isolate('גרסת Node.js 22 יצאה');
  assert.ok(out.includes('\u202DNode.js 22\u202C'));
  assert.ok(!out.startsWith('\u202D'));
});

test('isRtl() detects Hebrew but not Latin', () => {
  assert.ok(isRtl('שלום'));
  assert.ok(!isRtl('hello 123'));
});

test('paragraphs are bidirectional and runs are rightToLeft', async () => {
  const doc = new HebrewDoc();
  doc.add(p('שלום עולם'));
  const xml = await documentXml(doc);
  assert.ok(xml.includes('<w:bidi/>'), 'paragraph should carry w:bidi');
  assert.ok(xml.includes('<w:rtl/>'), 'run should carry w:rtl');
});

test('tables are marked visually right-to-left', async () => {
  const doc = new HebrewDoc();
  doc.add(table(['שם', 'כמות'], [['כרטיס', '2']]));
  const xml = await documentXml(doc);
  assert.ok(xml.includes('<w:bidiVisual/>'), 'table should carry w:bidiVisual');
});

test('numeric cells are centred and not reordered', async () => {
  const doc = new HebrewDoc();
  doc.add(table(['שם', 'טווח'], [['אפיק', '1 : 0..1']]));
  const xml = await documentXml(doc);
  assert.ok(xml.includes('\u202D1 : 0..1\u202C'), 'numeric cell should be direction-locked');
});

test('headings and bullets render without throwing', async () => {
  const doc = new HebrewDoc({ font: 'Arial', size: 11 });
  doc.add(heading('כותרת', 1), bullet('פריט ראשון'), p([run('טקסט '), run(ltr('42'))]));
  const xml = await documentXml(doc);
  assert.ok(xml.includes('כותרת'));
  assert.ok(xml.includes('כותרת'));
});

test('landscape section flips the page size', async () => {
  const doc = new HebrewDoc();
  doc.add(p('לאורך'));
  doc.section({ landscape: true });
  doc.add(p('לרוחב'));
  const xml = await documentXml(doc);
  assert.ok(xml.includes('w:orient="landscape"'));
});
