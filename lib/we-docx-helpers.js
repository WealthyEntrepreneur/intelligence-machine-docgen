// WE-CP-DOCX-HELPERS-V2.js
// Wealthy Entrepreneur - DOCX Helper Functions Library
// Version 2.3 - June 2026
//
//
// PATCH NOTES (V2.3, June 2026): Build hardening merge. Folds the proven scaffold
// fixes (Managed Hosting Recommendation build) into the canonical library so every
// future build inherits them instead of rediscovering them, per WE-DOCX-BUILD-
// HARDENING. Additive and backward compatible; every prior export keeps its name.
// (1) keepTogether(inner): wraps any table or box in a single cantSplit outer row so
// the whole unit moves together and cannot split between rows (setting cantSplit on
// inner rows alone does NOT stop a multi-row table from breaking). (2) The split-prone
// builders now route through keepTogether so none can split: box, calloutBox, quoteBox,
// statBox, tbl (the data table; pass {splittable:true} to allow a genuinely long table
// to break at a natural row boundary per Formatting Guidelines Rule 2), the new
// optionsTable (DO-54 Implementation Options Comparison Table), the new disclosure
// (DO-54 Honest Disclosure Block, locked wording), and closingBlock. (3) Indent 0 on
// every table and box so the border lands on the body-text left margin on both sides
// (the prior bug was a wrong-sign indent that pushed borders about 108 twips left;
// border alignment is verified in the render-verify preflight). (4) The closing block
// flows as a unit at the foot of the body; buildDocument places it in its OWN final
// section, centered by an exact-line top spacer, only when the caller passes o.closing
// (the alone-on-page case), avoiding the Rule 5 blank-page double. (5) gap() and
// spaced(arr): spaced() inserts outside breathing room (a gap paragraph) before and
// after every box and data table in a children array, keyed off a non-enumerable
// __box mark that keepTogether sets. (6) h1 now carries keepNext so a top-level heading
// cannot orphan at a page foot (heading-block binder, Guidelines 1.4). Naming is
// reconciled to the existing NAVY/TEAL/run constants; no duplicate color object or run
// builder is introduced. Filename kept at V2 so require() lines and the guidelines
// pointer do not move; prior file (V2.2) archives per the Self-Contained File Rule.
//
// PATCH NOTES (V2.2, June 2026): Native updatable Word table of contents,
// additive and backward compatible. New helper tableOfContents(title) returns a
// "Contents" title paragraph plus a native Word TOC field over heading levels 1 to 3;
// h1/h2/h3 now carry HeadingLevel.HEADING_1/2/3 so the field can find them (the WE run
// styling is unchanged, headings look identical and simply gain an outline level);
// buildDocument() sets features.updateFields so Word populates and refreshes the TOC on
// open; TableOfContents is imported from docx and re-exported alongside tableOfContents.
// Word-optimized: Google Docs does not render Word TOC fields cleanly, so use this when
// the deliverable home is Word, and keep the manual clickable-list approach for
// Google-bound documents. Filename kept at V2 so require() lines and the guidelines
// pointer do not move; prior file (V2.1) archives per the Self-Contained File Rule.
// PATCH NOTES (V2.1, June 2026): Additive client-deliverable layer, backward
// compatible. Every V2.0 export keeps its name and behavior, so existing callers
// (WE-CP-DOCX-USAGE-EXAMPLES, the questionnaire render.js, and the deliverable
// templates) render unchanged. p() and bullet() now ALSO accept an array of runs;
// their string behavior is byte-identical to V2.0. New, non-colliding additions
// fold in the client-deliverable capabilities and lock in WE-CP-DOCX-FORMATTING-
// GUIDELINES: setLogo/logoImg (requires a true-transparent mark), run/runB,
// section() (flowing Navy header with a Teal rule, bound to its content),
// calloutBox() (multi-paragraph box; box() stays for single-paragraph callouts),
// coverClientBranded(), closingBlock() (C6/C7 with the locked promise),
// runningHeader(), footerText()/footerFor() (per-class footers; the client class
// renders the two-column co-brand footer), numbered(), spacer(), pageBreak().
// buildDocument() is OVERLOADED: the V2.0 positional form
// buildDocument(children, practiceName, documentTitle) is unchanged; a new options
// form buildDocument({ cover, children, closing, footerClass, client, date,
// headerText, closingOpts }) builds client deliverables with the cover title page,
// per-class footer, and a centered closing block in its own section (Rule 2C). Both
// forms register the weBullets/weNumbers numbering. Filename kept at V2 so require()
// lines and the guidelines pointer do not move; prior file (V2.0) archives per the
// Self-Contained File Rule.
//
// PATCH NOTES (V2.0, June 2026): Brand correction. The retired blue/gold/Calibri
// palette (BLUE 1F4E79, GOLD B8860B, FONT Calibri) is replaced with the locked
// Navy/Teal/Arial system per WE-BRAND-DESIGN-SYSTEM V1.1. Functional colors moved
// to the brand set (Red B83230, Green 2A7D4F, Amber C8910A replacing the old gold,
// Slate 5A6A7A). Light-background pairs aligned to Section 2.2. Content width
// corrected from 9360 to 10080 to match the 0.75 inch margins and the on-brand
// questionnaire generator (render.js). The full function API and export names are
// unchanged. Legacy names BLUE, GOLD, GRAY, LIGHT_BLUE, LIGHT_GOLD are retained as
// aliases pointing at the brand values, so existing callers do not break. Prior
// version was V1.0 (February 2026); archive it per the Self-Contained File Rule.
//
// Usage: const helpers = require("./WE-CP-DOCX-HELPERS-V2.js");
// const { h1, h2, h3, p, pb, sp, box, tbl, statBox, quoteBox,
//   bullet, num, check, buildDocument, saveDocument } = helpers;

const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat, ImageRun, TableOfContents,
  VerticalAlign } = require("docx");

// ============================================================
// COLOR SYSTEM (WE-BRAND-DESIGN-SYSTEM V1.1)
// ============================================================
// Primary brand
const NAVY = "0C1C3B";   // Primary. Headlines, table headers, H1, buttons
const TEAL = "4393A5";   // Accent. H3, subheadings, highlight text, key insights
const MINT = "E8F4F4";   // Brand background
const WHITE = "FFFFFF";

// Functional / status (fixed, never tokenized)
const RED = "B83230";    // Critical issues, HIGH threat, urgent
const GREEN = "2A7D4F";  // Opportunities, positive outcomes, ON TRACK
const AMBER = "C8910A";  // Warnings, attention items (replaces the old gold)
const SLATE = "5A6A7A";  // Footers, secondary text, captions, metadata
const BODY = "333333";   // Body text, near-black

// Light backgrounds (boxes and containers), Section 2.2
const LIGHT_TEAL = "E8F4F4";  // Insight boxes, key takeaways
const LIGHT_NAVY = "EDF1F7";  // Quote boxes, testimonials, featured content
const LIGHT_RED = "FDECEC";   // Alert boxes, critical issues
const LIGHT_GREEN = "E8F5E9"; // Opportunity boxes, positive outcomes
const LIGHT_AMBER = "FEF5E7"; // Notice boxes, attention items, pre-filled tags
const LIGHT_GRAY = "F5F6F8";  // Stat boxes, data containers, neutral

// Legacy aliases (retained so existing callers do not break; values are brand-correct)
const BLUE = NAVY;
const GOLD = AMBER;
const GRAY = SLATE;
const LIGHT_BLUE = LIGHT_NAVY;
const LIGHT_GOLD = LIGHT_AMBER;

const LINE = "C7D0D9";   // Soft slate line for table and cell borders
const FONT = "Arial";    // DOCX font (Word + Google Docs cross-platform)
const CW = 10080;        // content width in DXA (US Letter, 0.75" margins: 12240 - 2*1080)
const PAGE = { width: 12240, height: 15840 };
const MARGIN = 1080;

// No-underscore aliases (used by the client-deliverable layer below)
const LIGHTNAVY = LIGHT_NAVY;
const LIGHTGRAY = LIGHT_GRAY;
const LIGHTRED = LIGHT_RED;
const LIGHTGREEN = LIGHT_GREEN;
const LIGHTAMBER = LIGHT_AMBER;

// ============================================================
// LOGO  (transparent navy/teal mark; set once per build) - V2.1
// ============================================================
let LOGO_BUF = null;
let LOGO_RATIO = 418 / 1277; // height / width of the WE mark
function setLogo(path, ratio) {
  LOGO_BUF = fs.readFileSync(path);
  if (ratio) LOGO_RATIO = ratio;
}
function logoImg(widthPx) {
  return new ImageRun({ type: "png", data: LOGO_BUF,
    transformation: { width: widthPx, height: Math.round(widthPx * LOGO_RATIO) } });
}

// Styled run builders used by the client-deliverable layer - V2.1
function run(text, o = {}) {
  return new TextRun({ text, font: FONT, size: o.size || 22, color: o.color || BODY,
    bold: o.bold || false, italics: o.italics || false });
}
function runB(text, o = {}) { return run(text, { ...o, bold: true }); }

// ============================================================
// BORDER PRESETS
// ============================================================
const border = { style: BorderStyle.SINGLE, size: 1, color: LINE };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: WHITE };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

// ============================================================
// BUILD-HARDENING CORE (V2.3)  -  see WE-DOCX-BUILD-HARDENING
// ============================================================

// Non-enumerable __box mark. spaced() reads it to add outside breathing room;
// it is non-enumerable so it never leaks into docx serialization.
function markBox(t) {
  Object.defineProperty(t, "__box", { value: true, enumerable: false, configurable: true });
  return t;
}

// keepTogether(inner): wrap any table or box in a single cantSplit outer row so the
// whole unit moves to the next page together and never splits between rows. Indent 0
// so the wrapper border lands on the text margin. Returns a __box-marked Table.
const _NILB = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder,
  insideHorizontal: noBorder, insideVertical: noBorder };
function keepTogether(inner) {
  const c = new TableCell({ width: { size: CW, type: WidthType.DXA }, borders: _NILB,
    margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [inner] });
  const outer = new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
    indent: { size: 0, type: WidthType.DXA }, borders: _NILB,
    rows: [new TableRow({ cantSplit: true, children: [c] })] });
  return markBox(outer);
}

// gap(): a thin breathing-room paragraph. spaced(arr): insert a gap before and after
// every __box element in a children array (no double gaps between adjacent boxes).
function gap() {
  return new Paragraph({ spacing: { before: 0, after: 0, line: 120 },
    children: [new TextRun({ text: " ", font: FONT, size: 8 })] });
}
function spaced(arr) {
  const out = []; let lastGap = false;
  for (const el of arr) {
    if (el && el.__box) {
      if (!lastGap) out.push(gap());
      out.push(el); out.push(gap()); lastGap = true;
    } else { out.push(el); lastGap = false; }
  }
  return out;
}

// ============================================================
// HEADING FUNCTIONS
// ============================================================

// H1 - Navy heading with bottom border
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    keepNext: true,
    spacing: { before: 360, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: NAVY } },
    children: [new TextRun({ text, font: FONT, size: 32, bold: true, color: NAVY })]
  });
}

// H2 - Navy subheading
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, font: FONT, size: 26, bold: true, color: NAVY })]
  });
}

// H3 - Teal sub-subheading (brand: H3 headings are Teal)
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, font: FONT, size: 22, bold: true, color: TEAL })]
  });
}

// ============================================================
// PARAGRAPH FUNCTIONS
// ============================================================

// Standard paragraph. A string is styled by opts (identical to V2.0); an array of
// runs is passed through. opts.line/after/align are optional extras (V2.1).
function p(content, opts = {}) {
  const spacing = { after: opts.after != null ? opts.after : 120 };
  if (opts.line) spacing.line = opts.line;
  return new Paragraph({
    spacing,
    alignment: opts.align,
    children: Array.isArray(content) ? content : [new TextRun({
      text: content,
      font: FONT,
      size: opts.size || 22,
      color: opts.color || BODY,
      bold: opts.bold || false,
      italics: opts.italics || false
    })]
  });
}

// Bold paragraph
function pBold(text) {
  return p(text, { bold: true });
}

// Paragraph with multiple styled runs
function pRuns(runs) {
  return new Paragraph({
    spacing: { after: 120 },
    children: runs.map(r => new TextRun({
      text: r.text,
      font: FONT,
      size: r.size || 22,
      color: r.color || BODY,
      bold: r.bold || false,
      italics: r.italics || false
    }))
  });
}

// Centered paragraph
function pCenter(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: opts.after || 120 },
    children: [new TextRun({
      text,
      font: FONT,
      size: opts.size || 22,
      color: opts.color || BODY,
      bold: opts.bold || false,
      italics: opts.italics || false
    })]
  });
}

// ============================================================
// SPACING FUNCTIONS
// ============================================================

// Page break
function pb() {
  return new Paragraph({ children: [new PageBreak()] });
}

// Spacer (small gap)
function sp() {
  return new Paragraph({ spacing: { after: 60 }, children: [] });
}

// ============================================================
// BOX FUNCTIONS
// ============================================================

// Colored insight/alert box with label.
// Pair per brand Section 2.2, e.g. box(text, "KEY TAKEAWAY", LIGHT_TEAL, TEAL).
function box(text, label, bgColor, labelColor) {
  return keepTogether(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    indent: { size: 0, type: WidthType.DXA },
    rows: [new TableRow({
      cantSplit: true,
      children: [new TableCell({
        borders: {
          top: { style: BorderStyle.SINGLE, size: 2, color: labelColor },
          bottom: noBorder, left: noBorder, right: noBorder
        },
        shading: { fill: bgColor, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        children: [
          new Paragraph({
            spacing: { after: 60 },
            children: [new TextRun({ text: label, font: FONT, size: 20, bold: true, color: labelColor })]
          }),
          new Paragraph({
            children: [new TextRun({ text, font: FONT, size: 22, color: BODY })]
          })
        ]
      })]
    })]
  }));
}

// Navy quote/testimonial box (brand: Light Navy fill, Navy rule)
function quoteBox(quote, source) {
  return keepTogether(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    indent: { size: 0, type: WidthType.DXA },
    rows: [new TableRow({
      cantSplit: true,
      children: [new TableCell({
        borders: {
          left: { style: BorderStyle.SINGLE, size: 4, color: NAVY },
          top: noBorder, bottom: noBorder, right: noBorder
        },
        shading: { fill: LIGHT_NAVY, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        children: [
          new Paragraph({
            spacing: { after: 60 },
            children: [new TextRun({ text: `"${quote}"`, font: FONT, size: 22, italics: true, color: BODY })]
          }),
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: `- ${source}`, font: FONT, size: 20, color: SLATE })]
          })
        ]
      })]
    })]
  }));
}

// Stat box row (3-4 metrics side by side). Brand: Light Gray fill, Navy value.
function statBox(stats) {
  const colWidth = Math.floor(CW / stats.length);
  const colWidths = stats.map(() => colWidth);
  // Adjust last column to absorb rounding
  colWidths[colWidths.length - 1] = CW - colWidth * (stats.length - 1);

  return keepTogether(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: colWidths,
    indent: { size: 0, type: WidthType.DXA },
    rows: [new TableRow({
      cantSplit: true,
      children: stats.map((stat, i) => new TableCell({
        borders,
        width: { size: colWidths[i], type: WidthType.DXA },
        shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 80, right: 80 },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 40 },
            children: [new TextRun({ text: String(stat.value), font: FONT, size: 32, bold: true, color: NAVY })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: stat.label, font: FONT, size: 18, color: SLATE })]
          })
        ]
      }))
    })]
  }));
}

// ============================================================
// TABLE FUNCTION
// ============================================================

// Table with navy header row, white header text. Wrapped through keepTogether by
// default so it never splits between rows. For a genuinely long table that must run
// past one page, pass opts.splittable = true to return the bare table, which then
// breaks at a natural row boundary per Formatting Guidelines Rule 2.
function tbl(headers, rows, colWidths, opts = {}) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      borders,
      width: { size: colWidths[i], type: WidthType.DXA },
      shading: { fill: NAVY, type: ShadingType.CLEAR },
      margins: cellMargins,
      children: [new Paragraph({
        children: [new TextRun({ text: h, font: FONT, size: 20, bold: true, color: WHITE })]
      })]
    }))
  });

  const dataRows = rows.map((row, ri) => new TableRow({
    cantSplit: true,
    children: row.map((cell, i) => new TableCell({
      borders,
      width: { size: colWidths[i], type: WidthType.DXA },
      shading: ri % 2 === 1 ? { fill: LIGHT_GRAY, type: ShadingType.CLEAR } : undefined,
      margins: cellMargins,
      children: [new Paragraph({
        children: [new TextRun({ text: String(cell), font: FONT, size: 20, color: BODY })]
      })]
    }))
  }));

  const table = new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: colWidths,
    indent: { size: 0, type: WidthType.DXA },
    rows: [headerRow, ...dataRows]
  });
  return opts.splittable ? table : keepTogether(table);
}

// ============================================================
// LIST FUNCTIONS
// ============================================================

// Note: These use simple text prefixes for compatibility.
// For proper Word bullet/number lists, use the numbering config
// approach shown in the DOCX skill documentation.

function bullet(content) {
  const body = Array.isArray(content) ? content
    : [new TextRun({ text: content, font: FONT, size: 22, color: BODY })];
  return new Paragraph({
    spacing: { after: 80 },
    indent: { left: 720 },
    children: [new TextRun({ text: "\u2022  ", font: FONT, size: 22, color: TEAL }), ...body]
  });
}

function num(text, number) {
  return new Paragraph({
    spacing: { after: 80 },
    indent: { left: 720 },
    children: [
      new TextRun({ text: `${number}.  `, font: FONT, size: 22, bold: true, color: NAVY }),
      new TextRun({ text, font: FONT, size: 22, color: BODY })
    ]
  });
}

function check(text, checked = false) {
  const symbol = checked ? "\u2611" : "\u2610";
  return new Paragraph({
    spacing: { after: 80 },
    indent: { left: 720 },
    children: [
      new TextRun({ text: `${symbol}  `, font: FONT, size: 22 }),
      new TextRun({ text, font: FONT, size: 22, color: BODY })
    ]
  });
}

// ============================================================
// CODE BLOCK
// ============================================================

function codeBlock(lines) {
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    indent: { size: 0, type: WidthType.DXA },
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 200, right: 200 },
        children: lines.map(line => new Paragraph({
          spacing: { after: 40 },
          children: [new TextRun({ text: line, font: "Consolas", size: 18, color: BODY })]
        }))
      })]
    })]
  });
}

// ============================================================
// CLIENT-DELIVERABLE LAYER (V2.1, additive)
// ============================================================

// Flowing section header: Navy with a Teal rule, bound to its content (keepNext,
// guidelines 1.4), NO forced page break. Use for DO-54 single Recommendations and
// General docs that flow with selective breaks. Use h1() where a class breaks every
// top-level section (Legal Articles, Bundle chapter titles).
function section(text) {
  return new Paragraph({
    spacing: { before: 300, after: 140 },
    outlineLevel: 0, keepNext: true,
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: TEAL, space: 4 } },
    children: [new TextRun({ text, font: FONT, size: 28, bold: true, color: NAVY })]
  });
}

// Multi-paragraph callout box: heading + array of paragraphs, neutral by default,
// navy left accent, never splits. Each para may be a string or an array of runs.
// (box() remains for single-paragraph, labeled callouts.)
function calloutBox(heading, paras, o = {}) {
  const fill = o.fill || LIGHT_GRAY;
  const accent = o.accent || NAVY;
  const txt = o.textColor || SLATE;
  const inner = [];
  if (heading) inner.push(new Paragraph({ spacing: { after: 80 },
    children: [new TextRun({ text: heading, font: FONT, size: 19, bold: true, color: txt })] }));
  paras.forEach((para, i) => inner.push(new Paragraph({
    spacing: { after: i === paras.length - 1 ? 0 : 100, line: 256 },
    children: Array.isArray(para) ? para : [new TextRun({ text: para, font: FONT, size: 18, color: txt })] })));
  const cb = { style: BorderStyle.SINGLE, size: 4, color: "D5DAE0" };
  return keepTogether(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
    indent: { size: 0, type: WidthType.DXA },
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: CW, type: WidthType.DXA },
      borders: { top: cb, bottom: cb, right: cb, left: { style: BorderStyle.SINGLE, size: 18, color: accent } },
      shading: { fill, type: ShadingType.CLEAR },
      margins: { top: 160, bottom: 160, left: 200, right: 200 },
      children: inner })] })] }));
}

// Spacer with an optional gap, and a standalone selective page break (Rule 5).
function spacer(after = 60) { return new Paragraph({ spacing: { after }, children: [new TextRun("")] }); }
function pageBreak() { return new Paragraph({ children: [new PageBreak()] }); }

// Auto-numbered list item. Accepts a string or an array of runs. Uses the weNumbers
// numbering registered by buildDocument.
function numbered(content) {
  return new Paragraph({
    numbering: { reference: "weNumbers", level: 0 },
    spacing: { after: 120, line: 276 },
    children: Array.isArray(content) ? content : [new TextRun({ text: content, font: FONT, size: 22, color: BODY })]
  });
}

// Client-branded cover (Formatting Guidelines 1.7). Returns an array of paragraphs
// ending in a page break; pass as the `cover` option to buildDocument.
function coverClientBranded(o) {
  const c = [];
  c.push(new Paragraph({ spacing: { before: 300, after: 0 }, alignment: AlignmentType.CENTER,
    children: [run(o.clientLogoText || "[ Client logo ]", { size: 20, italics: true, color: SLATE })] }));
  c.push(new Paragraph({ spacing: { before: 4200, after: 120 }, alignment: AlignmentType.CENTER,
    children: [run(o.title, { size: 52, bold: true, color: NAVY })] }));
  if (o.subtitle) c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 },
    children: [run(o.subtitle, { size: 24, italics: true, color: TEAL })] }));
  c.push(new Paragraph({ spacing: { before: 560, after: 60 }, alignment: AlignmentType.CENTER,
    children: [run(o.date, { size: 22, color: SLATE })] }));
  c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
    children: [run("Prepared for: " + (o.client || "[Client Name]"), { size: 22, color: SLATE })] }));
  c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 },
    children: [run("Prepared by Wealthy Entrepreneur     \u00B7     Confidential", { size: 20, color: SLATE })] }));
  c.push(new Paragraph({ spacing: { before: 2500, after: 0 }, alignment: AlignmentType.CENTER, children: [logoImg(235)] }));
  c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 90, after: 0 }, children: [
    run("Your edge", { size: 24, bold: true, color: TEAL }),
    run(" to grow bigger, move faster, and make a greater impact.", { size: 24, color: NAVY }) ] }));
  c.push(new Paragraph({ children: [new PageBreak()] }));
  return c;
}

// Standardized closing block (C6/C7): three-row table, tagline + resource row with
// the locked promise ("Your edge" in Teal, the rest in Navy).
function closingBlock(o = {}) {
  const url = o.url || "WealthyEntrepreneur.com";
  const tagLine = (text, last) => new Paragraph({ alignment: AlignmentType.CENTER,
    spacing: { after: last ? 0 : 40 }, children: [run(text, { size: 22, bold: true, color: NAVY })] });
  const none = { style: BorderStyle.NONE, size: 0, color: WHITE };
  const teal2 = { style: BorderStyle.SINGLE, size: 16, color: TEAL };
  const tealL = { style: BorderStyle.SINGLE, size: 18, color: TEAL };
  return keepTogether(new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [CW], indent: { size: 0, type: WidthType.DXA }, rows: [
    new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: CW, type: WidthType.DXA },
      borders: { top: none, bottom: none, right: none, left: tealL },
      shading: { fill: LIGHT_NAVY, type: ShadingType.CLEAR },
      margins: { top: 160, bottom: 160, left: 220, right: 200 },
      children: [ tagLine("Clarity creates Alignment."), tagLine("Alignment builds Momentum."),
        tagLine("Momentum produces Wealth.", true) ] })] }),
    new TableRow({ children: [new TableCell({
      width: { size: CW, type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: none },
      margins: { top: 20, bottom: 20, left: 0, right: 0 },
      children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: "", size: 8 })] })] })] }),
    new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: CW, type: WidthType.DXA },
      borders: { top: teal2, bottom: teal2, left: teal2, right: teal2 },
      shading: { fill: LIGHT_NAVY, type: ShadingType.CLEAR },
      margins: { top: 160, bottom: 160, left: 200, right: 200 },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [logoImg(205)] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
          children: [run(url, { size: 26, bold: true, color: TEAL })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [
          run("Your edge", { size: 22, bold: true, color: TEAL }),
          run(" to grow bigger, move faster, and make a greater impact.", { size: 22, color: NAVY }) ] }),
      ] })] }),
  ] }));
}

// Running header (all pages except cover).
function runningHeader(text) {
  return new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0 },
    children: [run(text, { size: 18, color: SLATE })] })] });
}

// Per-class footer text (1.9). Separator " | ", date "Month Day, Year".
function footerText(footerClass, o) {
  const date = o.date || "";
  const pageRun = [ run("Page ", { size: 18, color: SLATE }),
    new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: SLATE }),
    run(" of ", { size: 18, color: SLATE }),
    new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: SLATE }) ];
  const seg = (s) => run(s, { size: 18, color: SLATE });
  switch (footerClass) {
    case "client":
      return [ seg("Confidential. Prepared for " + (o.client || "[Client]") +
        ". Not for redistribution without Wealthy Entrepreneur consent.     |     " + date + "     |     "), ...pageRun ];
    case "internal-client":
      return [ seg((o.docName || "[Document]") + "     |     INTERNAL (WE TEAM ONLY)     |     Client: " +
        (o.client || "[Client]") + "     |     " + date + "     |     "), ...pageRun ];
    case "internal-we":
      return [ seg((o.docName || "[Document]") + "     |     INTERNAL (WE TEAM ONLY)     |     Company: " +
        (o.weEntity || "Wealthy Entrepreneur") + "     |     " + date + "     |     "), ...pageRun ];
    case "sop":
      return [ seg((o.docName || "[SOP Name]") + "     |     Internal SOP     |     v" + (o.version || "1") +
        "     |     " + date + "     |     "), ...pageRun ];
    case "legal-executed":
      return [ seg((o.docName || "[Document]") + "     |     "), ...pageRun ];
    default:
      return [ ...pageRun ];
  }
}

// Per-class footer. The client class renders the two-column co-brand footer (logo
// left, text stacked and right-aligned, both vertically centered) when a logo is set.
// All other classes render a single centered line.
function footerFor(footerClass, o = {}) {
  if (footerClass === "client" && LOGO_BUF) {
    const date = o.date || "";
    const line1 = "Confidential. Prepared for " + (o.client || "[Client]") + ". Not for redistribution without Wealthy Entrepreneur consent.";
    const nb = { style: BorderStyle.NONE, size: 0, color: WHITE };
    const nbAll = { top: nb, bottom: nb, left: nb, right: nb, insideHorizontal: nb, insideVertical: nb };
    const logoW = 1950, textW = CW - logoW;
    const tbl2 = new Table({
      width: { size: CW, type: WidthType.DXA }, columnWidths: [logoW, textW], borders: nbAll,
      indent: { size: 0, type: WidthType.DXA },
      rows: [new TableRow({ children: [
        new TableCell({ width: { size: logoW, type: WidthType.DXA }, borders: nbAll,
          verticalAlign: "center", margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 0 }, children: [logoImg(112)] })] }),
        new TableCell({ width: { size: textW, type: WidthType.DXA }, borders: nbAll,
          verticalAlign: "center", margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [
            new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 24 }, children: [run(line1, { size: 16, color: SLATE })] }),
            new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0 }, children: [
              run(date + "     |     Page ", { size: 16, color: SLATE }),
              new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: SLATE }),
              run(" of ", { size: 16, color: SLATE }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: SLATE }),
            ] }),
          ] }),
      ] })],
    });
    return new Footer({ children: [tbl2] });
  }
  return new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40 },
    children: footerText(footerClass, o) })] });
}

// Numbering config registered by both buildDocument forms (enables numbered()).
const NUMBERING = { config: [
  { reference: "weBullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022",
    alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 280 } } } }] },
  { reference: "weNumbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.",
    alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 280 } } } }] },
] };

// ============================================================
// DOCUMENT BUILDER
// ============================================================

// Overloaded. V2.0 positional form: buildDocument(children, practiceName, documentTitle)
// New options form: buildDocument({ cover, children, closing, footerClass, client,
//   date, headerText, closingOpts }) for client deliverables.
function buildDocument(arg1, practiceName, documentTitle) {
  if (Array.isArray(arg1)) {
    const children = arg1;
    const headerText = practiceName && documentTitle
      ? `${practiceName} | ${documentTitle}`
      : "Wealthy Entrepreneur";

    return new Document({
      numbering: NUMBERING,
      sections: [{
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
          }
        },
        headers: {
          default: new Header({
            children: [new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({
                text: headerText,
                font: FONT, size: 18, italics: true, color: SLATE
              })]
            })]
          })
        },
        footers: {
          default: new Footer({
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "Page ", font: FONT, size: 18, color: SLATE }),
                new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: SLATE }),
                new TextRun({ text: " | Wealthy Entrepreneur LLC | Confidential", font: FONT, size: 18, color: SLATE })
              ]
            })]
          })
        },
        children
      }]
    });
  }

  // Options-object form (client deliverables)
  const o = arg1 || {};
  const mainChildren = [];
  if (o.cover) mainChildren.push(...o.cover);
  mainChildren.push(...(o.children || []));

  const pageProps = { size: PAGE, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } };
  const emptyHdr = () => new Header({ children: [new Paragraph({ children: [new TextRun("")] })] });
  const emptyFtr = () => new Footer({ children: [new Paragraph({ children: [new TextRun("")] })] });
  const mkHeader = () => o.headerText ? runningHeader(o.headerText) : emptyHdr();
  const mkFooter = () => footerFor(o.footerClass || "client", o);

  const sections = [{
    properties: { titlePage: !!o.cover, page: pageProps },
    headers: { default: mkHeader(), first: emptyHdr() },
    footers: { default: mkFooter(), first: emptyFtr() },
    children: mainChildren,
  }];

  // Closing block in its own section, centered via a fixed-height spacer (Rule 2C).
  // The section break (not a manual page break) starts the page, avoiding the
  // blank-page double of Rule 5.
  if (o.closing) {
    sections.push({
      properties: { page: pageProps },
      headers: { default: mkHeader() },
      footers: { default: mkFooter() },
      children: [
        new Paragraph({ spacing: { line: 4200, lineRule: "exact" }, children: [new TextRun({ text: "\u00A0", font: FONT, size: 22 })] }),
        closingBlock(o.closingOpts || {}),
      ],
    });
  }

  return new Document({
    features: { updateFields: true },
    numbering: NUMBERING,
    styles: { default: { document: { run: { font: FONT, size: 22, color: BODY } } } },
    sections,
  });
}

// Save document to file
async function saveDocument(doc, filepath) {
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(filepath, buffer);
  console.log(`Created: ${filepath}`);
  return filepath;
}

// ============================================================
// TABLE OF CONTENTS (V2.2) - native updatable Word field
// ============================================================
// Returns [ "Contents" title paragraph, TOC field ]. The field lists h1/h2/h3
// (which now carry real Word heading styles). Word builds and updates it on open
// (buildDocument sets updateFields), and the reader refreshes it any time via
// right-click then Update Field. Place it on its own page (follow with pageBreak()).
// Word-optimized; Google Docs does not render Word TOC fields cleanly.
function tableOfContents(title) {
  return [
    new Paragraph({ spacing: { after: 120 },
      children: [new TextRun({ text: title || "Contents", font: FONT, size: 32, bold: true, color: NAVY })] }),
    new TableOfContents(title || "Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
  ];
}

// ============================================================
// DO-54 CLIENT-ADVISORY COMPONENTS (V2.3)  -  Formatting Guidelines 2.2
// ============================================================

// Implementation Options Comparison Table. Header light gray (#F5F6F8), navy text;
// the recommended row carries a Teal 3pt left border on its first cell only (no full-
// row highlight, so the honest Trade-offs content carries the persuasion). WE is named
// openly as the AI-augmented DFY path. Default headers follow Formatting Guidelines 2.2
// (Option, Implementation Path, Timeline, Investment, Trade-offs); pass a headers array
// to override. Routes through keepTogether so it never splits.
function optionsTable(rows, colWidths, recommendedIndex, headers) {
  const head = (headers || ["Option", "Implementation Path", "Timeline", "Investment", "Trade-offs"]);
  const headerRow = new TableRow({
    tableHeader: true,
    children: head.map((h, i) => new TableCell({
      borders,
      width: { size: colWidths[i], type: WidthType.DXA },
      shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR },
      margins: cellMargins, verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ children: [new TextRun({ text: h, font: FONT, size: 20, bold: true, color: NAVY })] })]
    }))
  });
  const tealLeft = { style: BorderStyle.SINGLE, size: 24, color: TEAL };
  const dataRows = rows.map((row, ri) => {
    const fill = ri % 2 === 1 ? LIGHT_GRAY : WHITE;
    const isRec = ri === recommendedIndex;
    return new TableRow({ cantSplit: true, children: row.map((cell, ci) => new TableCell({
      borders: (ci === 0 && isRec) ? { ...borders, left: tealLeft } : borders,
      width: { size: colWidths[ci], type: WidthType.DXA },
      shading: { fill, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 140, right: 140 }, verticalAlign: VerticalAlign.CENTER,
      children: Array.isArray(cell) ? cell
        : [new Paragraph({ children: [new TextRun({ text: String(cell), font: FONT, size: 20, color: BODY })] })]
    })) });
  });
  return keepTogether(new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: colWidths,
    indent: { size: 0, type: WidthType.DXA }, rows: [headerRow, ...dataRows]
  }));
}

// DO-54 Honest Disclosure Block. Light gray background, Navy 2pt top border only,
// italic Navy body, bold "Disclosure:" prefix, single page (never splits). The wording
// is locked per DO-54 Section 13 and Recommendations QA Gate Hard Stop 3; the Standard
// variant is the default. Pass `text` to supply the Bundle or WE-Service-Retirement
// locked wording. Routes through keepTogether.
const DISCLOSURE_STANDARD = "Wealthy Entrepreneur provides the AI-augmented DFY option referenced above. We have a vested interest in your implementation choice. We have included our offering alongside the alternatives so you can evaluate paths with full context. This Recommendation is delivered as part of your engagement regardless of which option you select. We are equally available to support you through any of the paths above, including by introducing you to traditional agencies or in-house hires if that is the right fit.";
function disclosure(text) {
  return keepTogether(new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [CW], indent: { size: 0, type: WidthType.DXA },
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: CW, type: WidthType.DXA },
      borders: { top: { style: BorderStyle.SINGLE, size: 12, color: NAVY }, bottom: noBorder, left: noBorder, right: noBorder },
      shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR },
      margins: { top: 320, bottom: 320, left: 400, right: 400 },
      children: [new Paragraph({ spacing: { line: 288 }, children: [
        new TextRun({ text: "Disclosure: ", font: FONT, size: 22, bold: true, color: NAVY }),
        new TextRun({ text: text || DISCLOSURE_STANDARD, font: FONT, size: 22, italics: true, color: NAVY })
      ] })]
    })] })]
  }));
}

// ============================================================
// EXPORTS
// ============================================================
module.exports = {
  // Brand colors
  NAVY, TEAL, MINT, WHITE,
  RED, GREEN, AMBER, SLATE, BODY,
  LIGHT_TEAL, LIGHT_NAVY, LIGHT_RED, LIGHT_GREEN, LIGHT_AMBER, LIGHT_GRAY,
  LINE, FONT, CW, PAGE, MARGIN,

  // Legacy aliases (brand-correct values)
  BLUE, GOLD, GRAY, LIGHT_BLUE, LIGHT_GOLD,
  // No-underscore aliases (client-deliverable layer)
  LIGHTNAVY, LIGHTGRAY, LIGHTRED, LIGHTGREEN, LIGHTAMBER,

  // Borders
  border, borders, noBorder, noBorders, cellMargins,

  // Logo + run builders (V2.1)
  setLogo, logoImg, run, runB,

  // Headings
  h1, h2, h3, section,

  // Paragraphs
  p, pBold, pRuns, pCenter,

  // Spacing
  pb, pageBreak, sp, spacer,

  // Build-hardening core (V2.3)
  keepTogether, markBox, gap, spaced,

  // Boxes
  box, calloutBox, quoteBox, statBox,

  // Tables
  tbl,

  // DO-54 client-advisory components (V2.3)
  optionsTable, disclosure,

  // Lists
  bullet, num, numbered, check,

  // Code
  codeBlock,

  // Client-deliverable layer (V2.1)
  coverClientBranded, closingBlock, runningHeader, footerText, footerFor,

  // Table of contents (V2.2)
  tableOfContents,

  // Document
  buildDocument, saveDocument,

  // Re-export docx classes for advanced usage
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat, ImageRun, TableOfContents,
  VerticalAlign
};
