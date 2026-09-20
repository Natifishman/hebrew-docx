'use strict';

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType, TableLayoutType, VerticalAlign,
  LevelFormat, PageOrientation, Footer, Header, PageNumber, PageBreak, HeadingLevel,
} = require('docx');

const LRO = '\u202D'; // LEFT-TO-RIGHT OVERRIDE
const PDF = '\u202C'; // POP DIRECTIONAL FORMATTING

const DEFAULTS = {
  font: 'David',
  size: 12,          // points
  lineSpacing: 1.15,
  pageSize: 'A4',
  margin: 2,         // cm
};

const PAGE_SIZES = {
  A4: { width: 11906, height: 16838 },
  Letter: { width: 12240, height: 15840 },
};

const cm = (n) => Math.round(n * 567);
const pt = (n) => Math.round(n * 2);          // docx half-points
const px = (n) => Math.round(n * 15);         // dxa -> px at 96dpi

/** True when the string contains Hebrew, Arabic or other RTL letters. */
function isRtl(text) {
  return /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/.test(String(text));
}

/**
 * Wraps a string in a directional override so it always reads left-to-right.
 * Use for numbers, versions, ranges, IDs and Latin terms sitting inside Hebrew
 * text, where the bidi algorithm would otherwise reorder them ("1 : 0..1"
 * rendered as "0..1 : 1").
 */
function ltr(text) {
  return LRO + String(text) + PDF;
}

/** Splits a string so Latin/number segments keep their own direction. */
function isolate(text) {
  return String(text).replace(/([A-Za-z0-9][A-Za-z0-9 .,:_\-+/()%]*[A-Za-z0-9)%]|[A-Za-z0-9])/g, (m) => ltr(m));
}

module.exports = {
  LRO, PDF, DEFAULTS, PAGE_SIZES, cm, pt, px, isRtl, ltr, isolate,
  docx: {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
    AlignmentType, BorderStyle, WidthType, ShadingType, TableLayoutType, VerticalAlign,
    LevelFormat, PageOrientation, Footer, Header, PageNumber, PageBreak, HeadingLevel,
  },
};

// The builder is attached afterwards so it can require this module's helpers.
Object.assign(module.exports, require('./builder'));
Object.assign(module.exports, require('./elements'));
