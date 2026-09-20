'use strict';

const fs = require('fs');
const {
  Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType, TableLayoutType, VerticalAlign,
  PageBreak,
} = require('docx');
const { DEFAULTS, cm, pt, ltr, isolate } = require('./index');

const ALIGN = {
  right: AlignmentType.RIGHT,
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  justify: AlignmentType.BOTH,
};

/**
 * A run of Hebrew text. `opts.ltr` forces left-to-right (for numbers or Latin),
 * `opts.isolate` protects embedded Latin/number segments inside Hebrew text.
 */
function run(text, opts = {}) {
  let value = String(text);
  if (opts.ltr) value = ltr(value);
  else if (opts.isolate) value = isolate(value);
  return new TextRun({
    text: value,
    rightToLeft: opts.ltr ? false : true,
    font: opts.font || DEFAULTS.font,
    size: pt(opts.size || DEFAULTS.size),
    bold: opts.bold,
    italics: opts.italic,
    underline: opts.underline ? {} : undefined,
    color: opts.color,
    break: opts.break,
  });
}

const toRuns = (content, opts) =>
  (Array.isArray(content) ? content : [content]).map((c) =>
    (c && typeof c === 'object' && c.constructor && c.constructor.name === 'TextRun') ? c : run(c, opts));

/** Right-aligned, bidirectional paragraph. */
function p(content, opts = {}) {
  return new Paragraph({
    bidirectional: true,
    alignment: ALIGN[opts.align] || AlignmentType.RIGHT,
    spacing: {
      before: opts.before != null ? pt(opts.before) * 10 : 0,
      after: opts.after != null ? pt(opts.after) * 10 : 120,
      line: Math.round((opts.lineSpacing || DEFAULTS.lineSpacing) * 240),
    },
    indent: opts.indent ? { right: cm(opts.indent) } : undefined,
    keepNext: opts.keepNext,
    children: toRuns(content, opts),
  });
}

/** Heading level 1-3. */
function heading(text, level = 1, opts = {}) {
  const sizes = { 1: 16, 2: 14, 3: 12.5 };
  return new Paragraph({
    heading: ['Heading1', 'Heading2', 'Heading3'][level - 1],
    bidirectional: true,
    alignment: ALIGN[opts.align] || AlignmentType.RIGHT,
    spacing: { before: 240, after: 120 },
    children: toRuns(text, { ...opts, bold: true, size: opts.size || sizes[level] }),
  });
}

/** Bulleted list item. Numbering is configured by the builder. */
const bullet = (content, opts = {}) => new Paragraph({
  bidirectional: true,
  alignment: AlignmentType.RIGHT,
  numbering: { reference: 'hb-bullet', level: opts.level || 0 },
  spacing: { after: 60, line: Math.round((opts.lineSpacing || DEFAULTS.lineSpacing) * 240) },
  children: toRuns(content, opts),
});

/** Numbered list item. Use a distinct `ref` to restart numbering. */
const numbered = (content, opts = {}) => new Paragraph({
  bidirectional: true,
  alignment: AlignmentType.RIGHT,
  numbering: { reference: opts.ref || 'hb-number', level: opts.level || 0 },
  spacing: { after: 60, line: Math.round((opts.lineSpacing || DEFAULTS.lineSpacing) * 240) },
  children: toRuns(content, opts),
});

const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

const cellParagraph = (text, opts) => {
  const value = String(text);
  const numeric = /^[\s\d.,:%+\-/()]*$/.test(value) && /\d/.test(value);
  return new Paragraph({
    bidirectional: !numeric,
    alignment: numeric ? AlignmentType.CENTER : (ALIGN[opts.align] || AlignmentType.RIGHT),
    spacing: { after: 40, line: 240 },
    children: [run(value, { ...opts, ltr: numeric })],
  });
};

/**
 * Table that reads right-to-left: the first column of `headers` is the
 * rightmost one on the page.
 *
 *   table(['שם', 'כמות'], [['כרטיס', '2']], { widths: [60, 40] })
 *
 * `widths` are percentages of the text width and default to equal columns.
 */
function table(headers, rows, opts = {}) {
  const cols = headers.length;
  const totalWidth = opts.width ? cm(opts.width) : 9638; // A4 minus 2cm margins
  const pct = opts.widths || Array(cols).fill(100 / cols);
  const widths = pct.map((w) => Math.round((w / 100) * totalWidth));
  const border = { style: BorderStyle.SINGLE, size: opts.borderSize || 4, color: opts.borderColor || '000000' };

  const cell = (content, i, isHeader) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    shading: isHeader && opts.headerFill
      ? { type: ShadingType.CLEAR, color: 'auto', fill: opts.headerFill } : undefined,
    children: Array.isArray(content) && content[0] instanceof Paragraph
      ? content
      : [cellParagraph(content, { size: opts.size, bold: isHeader })],
  });

  return new Table({
    visuallyRightToLeft: true,
    layout: TableLayoutType.FIXED,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border },
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, i, true)) }),
      ...rows.map((r) => new TableRow({ children: r.map((c, i) => cell(c, i, false)) })),
    ],
  });
}

function pngSize(buffer) {
  if (buffer.slice(1, 4).toString() !== 'PNG') return null;
  return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
}

/**
 * Centred image scaled to fit `maxWidth`/`maxHeight` (px), with an optional
 * caption paragraph underneath.
 */
function image(pathOrBuffer, opts = {}) {
  const data = Buffer.isBuffer(pathOrBuffer) ? pathOrBuffer : fs.readFileSync(pathOrBuffer);
  const type = opts.type || (String(pathOrBuffer).toLowerCase().endsWith('.jpg') ? 'jpg' : 'png');
  const size = pngSize(data) || { width: opts.width || 600, height: opts.height || 400 };
  const maxW = opts.maxWidth || 600;
  const maxH = opts.maxHeight || 800;
  const scale = Math.min(maxW / size.width, maxH / size.height, 1);
  const out = [new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: opts.caption ? 40 : 200 },
    keepNext: Boolean(opts.caption),
    children: [new ImageRun({
      type,
      data,
      transformation: { width: Math.round(size.width * scale), height: Math.round(size.height * scale) },
    })],
  })];
  if (opts.caption) {
    out.push(new Paragraph({
      bidirectional: true,
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [run(opts.caption, { size: opts.captionSize || 11 })],
    }));
  }
  return out;
}

module.exports = { run, p, heading, bullet, numbered, table, image, pageBreak, cellParagraph };
