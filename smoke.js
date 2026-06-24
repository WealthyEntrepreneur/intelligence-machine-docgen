// Local smoke test — renders a sample branded .docx (and PDF if LibreOffice is installed)
// without touching Supabase. Run: node smoke.js  -> writes ./out/smoke.docx (+ .pdf)
const fs = require('fs');
const path = require('path');
const { promisify } = require('util');
const { renderDocx, docxBuffer } = require('./lib/render');

const job = {
  tenant_id: 'demo', title: 'Strategic Recommendation',
  input: {
    client: 'TMG Plumbing & Disaster Solutions',
    title: 'Strategic Recommendation',
    subtitle: 'A sample branded deliverable',
    date: 'June 2026',
    blocks: [
      { type: 'stats', items: [
        { value: '380+', label: 'Reviews' }, { value: '4.8', label: 'Avg Rating' },
        { value: '92%', label: '5-Star' }, { value: '3', label: 'Locations' },
      ] },
      { type: 'h1', text: 'EXECUTIVE SUMMARY' },
      { type: 'p', text: 'This analysis synthesizes the latest intelligence for the client.' },
      { type: 'box', variant: 'teal', label: 'THE BOTTOM LINE', text: 'Strong fundamentals, with a visibility gap to close.' },
      { type: 'h1', text: 'RECOMMENDATIONS' },
      { type: 'numbered', items: ['Systematize review generation.', 'Lead marketing with patient language.'] },
      { type: 'table', headers: ['Platform', 'Reviews', 'Rating'], rows: [['Google', '285', '4.9'], ['Yelp', '52', '4.7']] },
      { type: 'quote', text: 'They made us feel like family.', source: 'Google Review' },
    ],
  },
};

(async () => {
  const out = path.join(__dirname, 'out');
  fs.mkdirSync(out, { recursive: true });
  const doc = renderDocx(job, {});
  const docx = await docxBuffer(doc);
  fs.writeFileSync(path.join(out, 'smoke.docx'), docx);
  console.log('Wrote out/smoke.docx (' + docx.length + ' bytes)');
  try {
    const libre = require('libreoffice-convert');
    const convertAsync = promisify(libre.convert);
    const pdf = await convertAsync(docx, '.pdf', undefined);
    fs.writeFileSync(path.join(out, 'smoke.pdf'), pdf);
    console.log('Wrote out/smoke.pdf (' + pdf.length + ' bytes)');
  } catch (e) {
    console.log('PDF step skipped (LibreOffice not installed locally): ' + e.message);
  }
})().catch((e) => { console.error(e); process.exit(1); });
