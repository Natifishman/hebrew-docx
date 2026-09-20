'use strict';

const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat,
  PageOrientation, Footer, Header, PageNumber,
} = require('docx');
const { DEFAULTS, PAGE_SIZES, cm, pt, ltr } = require('./index');

const numberingConfig = (font) => {
  const indent = (level) => ({
    paragraph: { indent: { left: 0, right: 720 + level * 540, hanging: 360 } },
  });
  const refs = ['hb-number', 'hb-number-2', 'hb-number-3', 'hb-number-4', 'hb-number-5'];
  return [
    {
      reference: 'hb-bullet',
      levels: [0, 1, 2].map((level) => ({
        level,
        format: LevelFormat.BULLET,
        text: ['•', '–', '·'][level],
        alignment: AlignmentType.RIGHT,
        style: { ...indent(level), run: { font } },
      })),
    },
    ...refs.map((reference) => ({
      reference,
      levels: [0, 1, 2].map((level) => ({
        level,
        format: LevelFormat.DECIMAL,
        text: level === 0 ? '%1.' : `%${level}.%${level + 1}`,
        alignment: AlignmentType.RIGHT,
        style: { ...indent(level), run: { font } },
      })),
    })),
  ];
};

/**
 * Collects RTL content and writes a .docx file.
 *
 *   const doc = new HebrewDoc({ font: 'David', size: 12 });
 *   doc.add(heading('כותרת'), p('שלום עולם'));
 *   await doc.save('out.docx');
 */
class HebrewDoc {
  constructor(options = {}) {
    this.options = { ...DEFAULTS, ...options };
    this.sections = [];
    this.current = { children: [], landscape: false };
  }

  /** Appends paragraphs, tables or arrays of them. */
  add(...items) {
    for (const item of items.flat(Infinity)) {
      if (item) this.current.children.push(item);
    }
    return this;
  }

  /**
   * Starts a new section, optionally in landscape - useful for wide tables
   * and diagrams inside an otherwise portrait document.
   */
  section({ landscape = false } = {}) {
    if (this.current.children.length) this.sections.push(this.current);
    this.current = { children: [], landscape };
    return this;
  }

  /** Page footer: "<text> - עמוד N". Applies to sections added afterwards. */
  footer(text) {
    this.options.footerText = text;
    return this;
  }

  _buildSection(section) {
    const { font, size, margin, pageSize, footerText } = this.options;
    const page = PAGE_SIZES[pageSize] || PAGE_SIZES.A4;
    const m = cm(section.landscape ? Math.max(margin - 0.75, 1) : margin);
    const properties = {
      page: {
        size: {
          ...page,
          orientation: section.landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT,
        },
        margin: { top: m, bottom: m, left: m, right: m },
      },
    };
    const footers = footerText ? {
      default: new Footer({
        children: [new Paragraph({
          bidirectional: true,
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: `${footerText} - עמוד `, rightToLeft: true, font, size: pt(size - 2) }),
            new TextRun({ children: [PageNumber.CURRENT], font, size: pt(size - 2) }),
          ],
        })],
      }),
    } : undefined;
    return { properties, footers, children: section.children };
  }

  /** Returns the underlying docx `Document`. */
  toDocument() {
    const { font, size } = this.options;
    const all = [...this.sections];
    if (this.current.children.length) all.push(this.current);
    if (!all.length) all.push({ children: [new Paragraph({})], landscape: false });

    const headingStyle = (id, name, points) => ({
      id, name, basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font, size: pt(points), bold: true, color: '000000' },
      paragraph: { outlineLevel: Number(id.slice(-1)) - 1 },
    });

    return new Document({
      styles: {
        default: { document: { run: { font, size: pt(size) } } },
        paragraphStyles: [
          headingStyle('Heading1', 'Heading 1', size + 4),
          headingStyle('Heading2', 'Heading 2', size + 2),
          headingStyle('Heading3', 'Heading 3', size + 0.5),
        ],
      },
      numbering: { config: numberingConfig(font) },
      sections: all.map((s) => this._buildSection(s)),
    });
  }

  /** Serialises to a Buffer. */
  toBuffer() {
    return Packer.toBuffer(this.toDocument());
  }

  /** Writes the document to disk. */
  async save(path) {
    fs.writeFileSync(path, await this.toBuffer());
    return path;
  }
}

module.exports = { HebrewDoc };
