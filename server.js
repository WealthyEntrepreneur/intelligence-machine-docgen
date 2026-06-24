// Intelligence Machine — document-generation service.
// POST /generate { job_id }  -> render branded .docx, convert to PDF, verify, upload, queue for review.
// POST /approve  { job_id }  -> mark approved (Drive delivery handled downstream on the 'approved' status).
// GET  /health
// Background poller: every POLL_MS, claim any 'queued' jobs and process them — so the app only has
// to INSERT a document_jobs row; no extra n8n wiring is required.
//
// Auth (HTTP): Authorization: Bearer <DOCGEN_TOKEN>.
const express = require('express');
const { promisify } = require('util');
const libre = require('libreoffice-convert');
const convertAsync = promisify(libre.convert);

const { supabase, getJob, getBrandKit, updateJob, uploadFile } = require('./lib/store');
const { renderDocx, docxBuffer } = require('./lib/render');
const { verify } = require('./lib/verify');
const { stripeCheckout, stripeWebhook } = require('./lib/billing');

const PORT = process.env.PORT || 8088;
const TOKEN = process.env.DOCGEN_TOKEN || '';
const POLL_MS = parseInt(process.env.POLL_MS || '8000', 10);

// Core pipeline: render -> PDF -> verify -> upload -> set status.
async function processJob(jobId) {
  const job = await getJob(jobId);
  await updateJob(jobId, { status: 'rendering', error: null });
  const brandKit = await getBrandKit(job.tenant_id);

  const doc = renderDocx(job, brandKit);
  const docx = await docxBuffer(doc);
  const pdf = await convertAsync(docx, '.pdf', undefined);

  const result = verify({ docxBuffer: docx, pdfBuffer: pdf, job });

  const stamp = Date.now();
  const safe = String(job.title || 'deliverable').replace(/[^a-zA-Z0-9._-]/g, '_').slice(0, 60);
  const base = `${job.tenant_id}/deliverables/${stamp}-${safe}`;
  const docxPath = await uploadFile(`${base}.docx`, docx,
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  const pdfPath = await uploadFile(`${base}.pdf`, pdf, 'application/pdf');

  await updateJob(jobId, {
    status: result.passed ? 'needs_review' : 'failed',
    checks: result.checks,
    docx_path: docxPath,
    pdf_path: pdfPath,
    error: result.passed ? null : 'render-verify failed',
  });
  return { status: result.passed ? 'needs_review' : 'failed', checks: result.checks, docx_path: docxPath, pdf_path: pdfPath };
}

const app = express();

// Stripe webhook needs the raw body for signature verification — register before express.json().
app.post('/billing/webhook', express.raw({ type: '*/*' }), stripeWebhook);

app.use(express.json({ limit: '4mb' }));

app.use((req, res, next) => {
  // Public/self-authenticating routes skip the shared-token gate.
  if (req.path === '/health' || req.path.startsWith('/billing/')) return next();
  if (!TOKEN || (req.headers.authorization || '') !== `Bearer ${TOKEN}`) return res.status(401).json({ error: 'unauthorized' });
  next();
});

app.get('/health', (_req, res) => res.json({ ok: true, service: 'im-docgen' }));

// Checkout authenticates the caller's Supabase JWT inside the handler.
app.post('/billing/checkout', stripeCheckout);

app.post('/generate', async (req, res) => {
  const jobId = req.body && req.body.job_id;
  if (!jobId) return res.status(400).json({ error: 'job_id required' });
  try {
    const out = await processJob(jobId);
    return res.json({ ok: true, ...out });
  } catch (e) {
    console.error('[docgen] generate failed', e);
    try { await updateJob(jobId, { status: 'failed', error: String(e.message || e) }); } catch (_) { /* ignore */ }
    return res.status(500).json({ error: String(e.message || e) });
  }
});

app.post('/approve', async (req, res) => {
  const jobId = req.body && req.body.job_id;
  if (!jobId) return res.status(400).json({ error: 'job_id required' });
  try {
    await updateJob(jobId, { status: 'approved', reviewed_by: (req.body && req.body.reviewed_by) || null });
    return res.json({ ok: true, status: 'approved' });
  } catch (e) {
    return res.status(500).json({ error: String(e.message || e) });
  }
});

// ---- Background poller ----
let polling = false;
async function poll() {
  if (polling) return;
  polling = true;
  try {
    const { data } = await supabase.from('document_jobs').select('id').eq('status', 'queued').order('created_at').limit(3);
    for (const j of (data || [])) {
      try { await processJob(j.id); } catch (e) {
        console.error('[docgen] poll job failed', j.id, e.message);
        try { await updateJob(j.id, { status: 'failed', error: String(e.message || e) }); } catch (_) { /* ignore */ }
      }
    }
  } catch (e) {
    console.error('[docgen] poll error', e.message);
  } finally {
    polling = false;
  }
}

app.listen(PORT, () => {
  console.log(`[docgen] listening on :${PORT}`);
  if (process.env.SUPABASE_URL) { setInterval(poll, POLL_MS); console.log(`[docgen] polling queued jobs every ${POLL_MS}ms`); }
});
