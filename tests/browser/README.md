# Hosted Chromium acceptance

This package keeps Playwright outside the application's dependency manifest and lockfile. Node 24.19.0 and npm 11.9.0 are the validated runtime. Playwright 1.64.0 is pinned exactly, along with its transitive dependency and registry integrity in this package's lockfile. The workflow audits this package separately and fails for any reported vulnerability.

The acceptance script is intended for a hosted GitHub Actions runner. It refuses to start outside Actions. It starts Vite on the runner's loopback interface, opens the actual app in Chromium, uploads only the included synthetic PDF fixtures, and receives genuine browser DOCX downloads. The fixed PDF.js CDN worker request is fulfilled from the installed, matching pdfjs-dist 3.11.174 worker. Any other external request, or any non-read request, is blocked and fails acceptance. Real dedicated PDF.js workers are required; fake-worker fallback fails acceptance.

## Coverage

- Healthy three-page scanned and text/vector PDFs, each with distinct red, green, and blue page markers
- Page 2 returning a null canvas context, an empty `data:,` export, malformed base64 that fails decoding, or an image-encoding exception
- For every fault: a page-specific error, no download, no success message, no 100% status, and no attempt to process page 3
- Retry on the same page with the same selected file after each fault, producing a complete DOCX
- Actual DOCX ZIP structure, exactly three media files, exact byte equality with all real source-page JPEGs in source order, page labels, title/date, and the existing image dimensions

Faults are limited to the browser canvas boundary on source page 2. PDF.js rendering, React, DOCX construction/packing, FileSaver, and downloads are not replaced. Source JPEG pixels also verify the expected page-color order. The independent core suite covers additional render/load/insert/pack/save exceptions and the complete progress-state sequence.

## Review evidence

The Actions artifact contains conversion screenshots, visible page text, each actual downloaded DOCX, each rendered source-page JPEG, DOCX/media hashes, browser logs, and the browser-tool audit report. `summary.json` records the actual checked-out commit and Git tree, event SHA, run ID/attempt, browser/runtime versions, both lockfile hashes, and source-fixture hashes. `manifest.json` binds each evidence file to its SHA-256 and that commit/tree.

Review the screenshots and render the downloaded DOCX files in a document viewer after downloading the artifact. The scripted ZIP/media checks do not claim that Word or LibreOffice layout has been visually accepted. Conversion acceptance is complete only after that separate visual review.

The PDF fixtures contain no user documents. Each is three pages at 480 × 240 points; the scanned version contains raster text markers and the vector version contains text and shapes. Both have page colors (220, 50, 50), (40, 150, 60), and (45, 85, 210), in that order. Their SHA-256 values are:

- `scanned-three-pages.pdf`: `3271e2e9dd398e37a12dbfd0d917f850a8d5c7fc59ab9a9c49f7f7e40f2b8661`
- `vector-three-pages.pdf`: `9fab680ced24e8b13a6ac2243c745238be0cfdda33ec0ccf86f47e1c2451d264`

## Existing dependency audit failures

The application lockfile is unchanged from baseline commit `e00efa9092aab809b40e1336a0b1672174ad47f2` (SHA-256 `5d5c0817fb159cc0444d0576c0f7d50009be75b77b96c0836b79a8acd1cfbe11`). The recorded baseline full audit fails with 31 findings: 1 low, 4 moderate, 25 high, and 1 critical. The production-only audit fails with 17 findings: 2 moderate, 14 high, and 1 critical. These are existing, unresolved findings; the rendering configuration mitigation does not make either audit pass.

The build, lint, conversion tests, and isolated browser-tool audit are separate checks. Their passing results must not be described as a passing application dependency audit. The existing full and production audit JSON/exit-code records are retained with the repair assessment. No runtime dependency upgrade or audit-gate suppression is part of this change.
