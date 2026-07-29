// R-40: Max-tier Deep Research also gets a 10-15 slide branded executive-summary deck.
// Gated by job.input.generate_pptx (set by the n8n execution workflow for deep_research_max only).
const PptxGenJS = require('pptxgenjs');

const MAX_CONTENT_SLIDES = 13; // + title slide + closing slide = 15 max (R-40)

function chunkParagraphIntoBullets(text, maxLen) {
  const sentences = String(text || '').split(/(?<=[.!?])\s+/).filter(Boolean);
  const bullets = [];
  let current = '';
  for (const s of sentences) {
    if ((current + ' ' + s).trim().length > maxLen && current) { bullets.push(current.trim()); current = s; }
    else { current = (current + ' ' + s).trim(); }
  }
  if (current) bullets.push(current.trim());
  return bullets;
}

async function pptxBuffer(job, brandKit) {
  const blocks = (job.input && job.input.blocks) || [];
  const input = job.input || {};
  const primary = ((brandKit && brandKit.primary_color) || '0C1C3B').replace('#', '');
  const accent = ((brandKit && brandKit.accent_color) || '4393A5').replace('#', '');
  const light = ((brandKit && brandKit.light_color) || 'FFFFFF').replace('#', '');

  const pptx = new PptxGenJS();
  pptx.defineSlideMaster({
    title: 'IM_MASTER',
    background: { color: light },
    objects: [{ rect: { x: 0, y: 0, w: '100%', h: 0.4, fill: { color: primary } } }],
  });

  const titleSlide = pptx.addSlide();
  titleSlide.background = { color: primary };
  titleSlide.addText(input.title || job.title || 'Deep Research', {
    x: 0.6, y: 2.2, w: 9, h: 1.4, fontSize: 32, bold: true, color: 'FFFFFF', fontFace: 'Arial',
  });
  titleSlide.addText(input.client || 'Client', {
    x: 0.6, y: 3.5, w: 9, h: 0.6, fontSize: 16, color: accent, fontFace: 'Arial',
  });
  if (input.date) {
    titleSlide.addText(String(input.date), { x: 0.6, y: 4.9, w: 9, h: 0.4, fontSize: 12, color: 'CCCCCC', fontFace: 'Arial' });
  }

  let contentSlides = 0;
  let currentSection = null;

  function newSectionSlide(heading) {
    const s = pptx.addSlide({ masterName: 'IM_MASTER' });
    if (heading) s.addText(heading, { x: 0.5, y: 0.7, w: 9, h: 0.7, fontSize: 24, bold: true, color: primary, fontFace: 'Arial' });
    s._bulletCount = 0;
    return s;
  }

  for (const b of blocks) {
    if (contentSlides >= MAX_CONTENT_SLIDES) break;
    const t = String(b.type || '').toLowerCase();
    if (t === 'h1') continue; // already on the title slide

    if (t === 'h2') {
      currentSection = newSectionSlide(b.text || '');
      contentSlides += 1;
      continue;
    }

    if (t === 'p' && b.text) {
      if (!currentSection) { currentSection = newSectionSlide(''); contentSlides += 1; }
      const bullets = chunkParagraphIntoBullets(b.text, 160);
      if (currentSection._bulletCount + bullets.length > 5 && contentSlides < MAX_CONTENT_SLIDES) {
        currentSection = newSectionSlide('');
        contentSlides += 1;
      }
      currentSection.addText(
        bullets.map((line) => ({ text: line, options: { bullet: true, breakLine: true } })),
        { x: 0.6, y: 1.6, w: 8.8, h: 4.8, fontSize: 16, color: '222222', fontFace: 'Arial', valign: 'top' },
      );
      currentSection._bulletCount += bullets.length;
      continue;
    }

    if (t === 'table' && Array.isArray(b.rows) && b.rows.length) {
      const s = newSectionSlide('');
      contentSlides += 1;
      const headers = (b.headers || []).map((h) => ({ text: String(h), options: { bold: true, color: 'FFFFFF', fill: { color: accent } } }));
      const rows = b.rows.slice(0, 8).map((r) => (r || []).map((c) => ({ text: String(c) })));
      s.addTable(headers.length ? [headers, ...rows] : rows, {
        x: 0.5, y: 1.6, w: 9, fontSize: 12, color: '222222', border: { type: 'solid', color: 'CCCCCC', pt: 0.5 },
      });
      currentSection = null;
    }
  }

  const closing = pptx.addSlide({ masterName: 'IM_MASTER' });
  closing.addText('Every report you buy, your machine keeps.', {
    x: 0.6, y: 2.6, w: 9, h: 1, fontSize: 20, italic: true, color: primary, fontFace: 'Arial', align: 'center',
  });

  return Buffer.from(await pptx.write({ outputType: 'nodebuffer' }));
}

module.exports = { pptxBuffer };
