# Intelligence Machine — Document Generation Service

Branded **.docx + PDF** generation for the Intelligence Machine portal. Ports the
`WE-CP-DOCX` helper library (Navy/Teal house style), converts to PDF with LibreOffice,
runs automated **render-verify** checks, stores both files in Supabase Storage, and
queues each deliverable for **human review** before delivery.

Runs as a small container next to your self-hosted n8n. **No extra n8n workflow is
required** — it polls `document_jobs` for `queued` rows and processes them.

## Flow

1. The app inserts a `document_jobs` row (`status = 'queued'`) with `input` (the document blocks).
2. This service claims it, renders `.docx`, converts to PDF, runs render-verify.
3. Pass → `status = 'needs_review'` (+ `docx_path`, `pdf_path`, `checks`). Fail → `status = 'failed'`.
4. An admin reviews in the portal and approves → `status = 'approved'`.
5. (Downstream) on `approved`, push the PDF to the client's Drive "Deliverables" folder. Download is already available from the stored paths.

## `input` shape

```json
{
  "client": "TMG Plumbing & Disaster Solutions",
  "title": "Strategic Recommendation",
  "subtitle": "A sample branded deliverable",
  "date": "June 2026",
  "blocks": [
    { "type": "stats", "items": [{ "value": "380+", "label": "Reviews" }] },
    { "type": "h1", "text": "EXECUTIVE SUMMARY" },
    { "type": "p", "text": "..." },
    { "type": "box", "variant": "teal", "label": "THE BOTTOM LINE", "text": "..." },
    { "type": "numbered", "items": ["Step one.", "Step two."] },
    { "type": "table", "headers": ["A", "B"], "rows": [["1", "2"]] },
    { "type": "quote", "text": "...", "source": "..." }
  ]
}
```
Block types: `h1 h2 h3 section p spacer pagebreak bullets numbered stats quote callout table box`.
Box `variant`: `teal insight red critical green opportunity amber notice navy quote`.

## Run

```bash
cp .env.example .env   # fill SUPABASE_SERVICE_ROLE_KEY + a long DOCGEN_TOKEN
docker build -t im-docgen .
docker run -d --name im-docgen --env-file .env -p 8088:8088 --restart unless-stopped im-docgen
curl localhost:8088/health
```

Put it on the same Docker network as n8n if you want n8n to reach it by name (`http://im-docgen:8088`).

## Local smoke test (no Supabase / no Docker)

```bash
npm install
node smoke.js   # writes out/smoke.docx (PDF only if LibreOffice is installed locally)
```

## Endpoints

- `GET /health` — liveness.
- `POST /generate { job_id }` — process one job now (the poller does this automatically). Needs `Authorization: Bearer <DOCGEN_TOKEN>`.
- `POST /approve { job_id, reviewed_by }` — mark a reviewed job approved.

## Notes

- The service-role key lives **only** in this container's env — never in the browser bundle.
- Fonts: the helpers use "Arial"; the image installs Liberation (metric-compatible) so PDFs match.
- Per-client brand kit (`brand_kits` table) themes the cover + co-brand footer. Full per-heading recolor is a future enhancement.
