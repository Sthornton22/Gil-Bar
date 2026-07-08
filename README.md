# Gil-Bar Proposal Engine (Render service)

Same process as the offline single-file builder, running as a web service:
drop PDFs -> classify (selection / cut sheet / drawing set) -> line-accurate
recognition (any manufacturer's Model/Qty/Tag/Voltage labels) -> design-basis
resolution + sister-family comparison trap -> scope-gate default-EXCLUDE ->
verbatim-precedent scoring -> styled Gil-Bar .docx (blank pricing, 20-clause
T&C, red [CONFIRM] for anything the documents don't contain). No pricing ever.

## Deploy to Render (replaces the current service)

1. Put these files in the repo connected to the `gilbar-proposal-generator`
   Render service (replace the existing contents).
2. Service type: **Web Service** · Runtime: **Node** (>= 18)
   - Build command: `npm install`
   - Start command: `npm start`
3. Push. Render redeploys automatically.

No environment variables, database, or disk needed. Drafts and history are
kept in memory only: they clear on restart/redeploy, and uploaded files are
processed per-request and never stored.

## Endpoints
- `GET /` — upload page (multi-file: PDF + Excel/CSV cost sheets, drag-drop, auto-classification)
- `POST /generate` — runs the pipeline, returns the review page + download link
- `GET /download/:id` — the generated .docx
- `GET /history` — drafts from this running instance

## Layout
- `server.js` — web service + pipeline orchestration
- `lib/app_core.js` — the verified process core (identical file used by the
  offline builder; do not fork the logic — fix it here and copy to both)
- `assets/` — baked workbook payload, verbatim 20-clause T&C, wordmark

## Honest limits
- A scrambled/image-only AutoCAD text layer cannot be read as text; the result
  page flags it (OCR would be needed).
- The letterhead is the reconstructed builder-kit letterhead, not the
  embedded-wordmark master binary.

## Cost sheet handling
- Excel/CSV (or PDF) cost sheets are the source of truth for the quoted set:
  models, tags, and quantities come from the sheet.
- Pricing columns (COST/PRICE/SELL/NET/TOTAL/$) are detected and SKIPPED —
  their values are never read into memory and never appear anywhere.
- A selection package alongside the cost sheet corroborates voltage and exposes
  the sister-family comparison side; qty/voltage mismatches are flagged.
- Cost-sheet option text is NOT converted into Includes — includes still come
  from the banked verbatim library with a confirm flag (verbatim rule).
- A by-others line found on a cost sheet is left off the proposal and flagged.
