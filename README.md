# hebrew-docx
Generate Word (`.docx`) documents in Hebrew — or any right-to-left language —
from Node.js, with correct bidirectional text, right-to-left tables, and page
layout that doesn't need fixing by hand afterwards.

<p align="center">
  <img src="assets/banner.png" alt="hebrew-docx" width="320">
</p>

## The problem

[docx](https://www.npmjs.com/package/docx) is a solid low-level library for
building `.docx` files, but it has no opinion about RTL. Left to its
defaults, Hebrew content built with it breaks in predictable ways:

- Paragraphs render left-aligned unless every one is marked bidirectional.
- Tables read left-to-right, so the "first" column ends up on the wrong side.
- A number or a Latin term sitting inside a Hebrew sentence gets reordered by
  the bidi algorithm — `"1 : 0..1"` can render as `"0..1 : 1"`, and nobody
  notices until the document is open in Word.
- Numbered and bulleted lists need their own RTL-aware numbering
  definitions, or the markers end up on the left.

None of this is a docx-js bug — it's just BiDi text layout, which is genuinely
fiddly to get right from scratch every time. `hebrew-docx` bakes in the fixes
once so a generated document reads correctly by default.

## Install

```bash
npm install hebrew-docx
```

## Quick example

```js
const { HebrewDoc, heading, p, bullet, table, run, ltr } = require('hebrew-docx');

const doc = new HebrewDoc({ font: 'David', size: 12 });
doc.footer('דוח לדוגמה');

doc.add(
  heading('דוח מכירות', 1),
  p('המכירות ברבעון השלישי עלו לעומת הרבעון הקודם.'),
  bullet('צפון: עלייה של 12%'),
  bullet(['דרום: עלייה של ', run(ltr('8.4%'))]),
  table(
    ['אזור', 'מכירות'],
    [['צפון', '₪120,000'], ['דרום', '₪98,500']],
    { widths: [60, 40], headerFill: 'D9D9D9' },
  ),
);

doc.save('report.docx');
```

The result opens in Word with the title right-aligned, the bullets reading
right-to-left, the table's first column ("אזור") on the right, and the
percentage rendering as `8.4%` rather than `%4.8`.

Full runnable examples: [`examples/invoice.js`](examples/invoice.js) (a
Hebrew invoice with a totals table) and [`examples/report.js`](examples/report.js)
(headings, lists, and a landscape section for a wide table).

## Why `ltr()` and `isolate()`

Hebrew text flows right-to-left, but numbers, dates, and Latin words inside
it are still left-to-right internally. The Unicode bidi algorithm handles
simple cases on its own, but ranges, ratios, and code-like strings
(`"1 : 0..1"`, `"Node.js 18 → 22"`) are exactly the cases it gets wrong.

- **`ltr(text)`** wraps a string in a directional override
  (`U+202D … U+202C`), locking its internal order. Use it for a whole
  value that must never be reordered — a version string, a date, a ratio.
- **`isolate(text)`** scans a mixed Hebrew/Latin sentence and wraps only the
  Latin/numeric segments, leaving the Hebrew words alone. Use it for prose
  that quotes an identifier or a number in passing.

Table cells that are purely numeric are detected and direction-locked
automatically, so `table()` output usually needs no manual handling.

## API

### `new HebrewDoc(options)`

| option | default | description |
|---|---|---|
| `font` | `'David'` | Default font for the whole document |
| `size` | `12` | Default font size in points |
| `lineSpacing` | `1.15` | Line spacing multiplier |
| `pageSize` | `'A4'` | `'A4'` or `'Letter'` |
| `margin` | `2` | Page margin in cm |

- **`.add(...items)`** — appends paragraphs, tables, or arrays of them (so
  `image()`, which returns multiple paragraphs, can be spread directly in).
- **`.section({ landscape })`** — starts a new section; content added after
  this call goes on new pages, optionally landscape. Useful for a wide
  table or diagram in an otherwise portrait report.
- **`.footer(text)`** — sets a centred footer, `"<text> - עמוד N"`.
- **`.save(path)`** — writes the `.docx` file, returns the path.
- **`.toBuffer()`** / **`.toDocument()`** — for custom output handling, or
  to combine with the raw `docx` API for anything this library doesn't
  cover yet.

### Elements (`require('hebrew-docx')`)

| function | purpose |
|---|---|
| `p(content, opts)` | Right-aligned, bidirectional paragraph |
| `heading(text, level, opts)` | Heading, level 1–3 |
| `bullet(content, opts)` | Bulleted list item (RTL markers) |
| `numbered(content, opts)` | Numbered list item; pass `ref` to restart a sequence |
| `table(headers, rows, opts)` | RTL table — `headers[0]` is the rightmost column |
| `image(path, opts)` | Centred, scaled image with an optional caption |
| `run(text, opts)` | A single text run; `opts.ltr` / `opts.isolate` control direction |
| `ltr(text)` / `isolate(text)` | Direction-locking helpers, usable outside `run()` too |
| `isRtl(text)` | `true` if the string contains Hebrew/Arabic characters |
| `pageBreak()` | Explicit page break |

`opts` on paragraph-like elements accepts `align` (`'right' | 'left' | 'center' | 'justify'`),
`bold`, `italic`, `underline`, `color`, `size`, `font`.

`table()` options: `widths` (percentages per column), `width` (total width
in cm), `headerFill` (hex color), `borderSize`, `borderColor`.

This library re-exports the full `docx` API under `docx` (`const { docx } =
require('hebrew-docx')`), so anything not covered here — headers, footnotes,
comments, sections with different page numbering — is still reachable
without dropping down to a separate import.

## Testing

```bash
npm test
```

Tests unzip the generated `.docx` and assert on the underlying
`word/document.xml` — the RTL flags (`w:bidi`, `w:rtl`, `w:bidiVisual`) and
the directional overrides land where they should, not just that the
document opens without throwing.

## Related

[rtl-diagram](https://github.com/Natifishman/rtl-diagram) — a small Python
engine for laying out diagrams (flowcharts, DFDs, entity-relationship
diagrams) with correct Hebrew/RTL text, exporting to SVG and to editable
draw.io files. Built for the same class of problem: RTL layout that
general-purpose tools don't handle out of the box.

## License

MIT
