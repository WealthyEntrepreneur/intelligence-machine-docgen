// Automated render-verify: cheap structural checks before a human review.
// Anything that fails routes the job to 'failed'; anything that passes routes to 'needs_review'.
const SSN_RE = /\b\d{3}-\d{2}-\d{4}\b/;

function verify({ docxBuffer, pdfBuffer, job }) {
  const input = job.input || {};
  const blocks = input.blocks || [];
  const checks = [];
  const add = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail || '' });

  add('docx_nonempty', docxBuffer && docxBuffer.length > 1024, `${docxBuffer ? docxBuffer.length : 0} bytes`);
  add('pdf_nonempty', pdfBuffer && pdfBuffer.length > 1024, `${pdfBuffer ? pdfBuffer.length : 0} bytes`);
  add('pdf_valid_header', pdfBuffer && pdfBuffer.slice(0, 5).toString('latin1') === '%PDF-', 'PDF magic bytes');
  add('has_title', !!(input.title || job.title), input.title || job.title || '(none)');
  add('has_content', blocks.length > 0, `${blocks.length} blocks`);

  // HIPAA gate: no patient-identifying SSNs in the rendered content.
  const flat = JSON.stringify(input);
  add('no_phi_ssn', !SSN_RE.test(flat), SSN_RE.test(flat) ? 'SSN-like pattern found' : 'clean');

  const passed = checks.every((c) => c.ok);
  return { passed, checks };
}

module.exports = { verify };
