// Turn a job's structured `input` into a branded docx Document using the WE-CP-DOCX helpers.
const path = require('path');
const H = require('./we-docx-helpers.js');

// Register the WE co-brand mark for the footer (shipped beside this file).
try { H.setLogo(path.join(__dirname, 'we-logo.png'), 418 / 1277); } catch (e) { /* footer falls back to text */ }

const BOX_VARIANTS = {
  teal: ['LIGHT_TEAL', 'TEAL'], insight: ['LIGHT_TEAL', 'TEAL'], takeaway: ['LIGHT_TEAL', 'TEAL'],
  red: ['LIGHT_RED', 'RED'], critical: ['LIGHT_RED', 'RED'], priority: ['LIGHT_RED', 'RED'],
  green: ['LIGHT_GREEN', 'GREEN'], opportunity: ['LIGHT_GREEN', 'GREEN'],
  amber: ['LIGHT_AMBER', 'AMBER'], notice: ['LIGHT_AMBER', 'AMBER'], warning: ['LIGHT_AMBER', 'AMBER'],
  navy: ['LIGHT_NAVY', 'NAVY'], quote: ['LIGHT_NAVY', 'NAVY'],
};

function renderBlock(b) {
  if (!b || !b.type) return [];
  const t = String(b.type).toLowerCase();
  switch (t) {
    case 'h1': return [H.h1(b.text || '')];
    case 'h2': return [H.h2(b.text || '')];
    case 'h3': return [H.h3(b.text || '')];
    case 'section': return [H.section(b.text || '')];
    case 'p': case 'paragraph': return [H.p(b.text || '')];
    case 'spacer': case 'sp': return [H.sp()];
    case 'pagebreak': case 'pb': return [H.pb()];
    case 'bullets': return (b.items || []).map((i) => H.bullet(String(i)));
    case 'numbered': case 'steps': return (b.items || []).map((i) => H.numbered(String(i)));
    case 'stats': return [H.statBox((b.items || []).map((s) => ({ value: String(s.value ?? ''), label: String(s.label ?? '') })))];
    case 'quote': return [H.quoteBox(b.text || '', b.source || '')];
    case 'callout': return [H.calloutBox(b.title || '', Array.isArray(b.paras) ? b.paras.map(String) : [String(b.text || '')])];
    case 'table': {
      const headers = b.headers || [];
      const rows = (b.rows || []).map((r) => r.map(String));
      const n = Math.max(headers.length, 1);
      // Default to equal columns across the content width when widths aren't supplied.
      const widths = Array.isArray(b.widths) && b.widths.length === n
        ? b.widths
        : Array.from({ length: n }, () => Math.floor(H.CW / n));
      return [H.tbl(headers, rows, widths)];
    }
    case 'box': {
      const [bg, lc] = BOX_VARIANTS[String(b.variant || 'teal').toLowerCase()] || BOX_VARIANTS.teal;
      return [H.box(b.text || '', b.label || '', H[bg], H[lc])];
    }
    default:
      return b.text ? [H.p(String(b.text))] : [];
  }
}

// input: { client, title, subtitle, date, headerText, blocks: [...] }
// logoBuf: optional image Buffer (the client's brand-kit logo), embedded on the cover.
function renderDocx(job, brandKit, logoBuf) {
  // Brand the whole document in the client's color tokens (never WE Navy/Teal for client work).
  if (brandKit && (brandKit.primary_color || brandKit.accent_color)) {
    H.setBrand({ primary: brandKit.primary_color, accent: brandKit.accent_color });
  } else {
    H.resetBrand();
  }
  const input = job.input || {};
  const client = input.client || 'Client';
  const title = input.title || job.title || 'Deliverable';
  const date = input.date || new Date().toISOString().slice(0, 10);

  const children = [];
  for (const b of (input.blocks || [])) children.push(...renderBlock(b));
  if (!children.length) children.push(H.p('This document has no content yet.'));

  return H.buildDocument({
    cover: H.coverClientBranded({
      title,
      subtitle: input.subtitle || '',
      date,
      client,
      logo: logoBuf || null,
    }),
    children,
    closing: true,
    footerClass: 'client',
    client,
    date,
    headerText: input.headerText || `${client} - ${title}`,
  });
}

async function docxBuffer(doc) {
  return H.Packer.toBuffer(doc);
}

module.exports = { renderDocx, docxBuffer };
