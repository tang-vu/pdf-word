# Hosted production and development acceptance

Playwright stays outside the application's dependency manifest and lockfile. The existing isolated package pins Playwright 1.64.0; both application and tooling locks remain unchanged. The workflow uses Node 24.19.0 and npm 11.9.0, audits the isolated tooling with an enforced zero-findings gate, and preserves its JSON report and actual exit code.

The script requires both genuine GitHub Actions and a GitHub-hosted runner. Do not spoof those flags or launch this suite locally. It installs and runs the Chromium, Firefox and WebKit engines belonging to the pinned Playwright release. These current hosted engines do not establish support for Safari, iOS or older browser versions. A configured engine is accepted only when its results are present and pass in the artifact.

## Served production worker

The suite runs the normal `npm run build` script separately for `/` and `/pdf-word/`, including normal minification and the application's `esbuild.keepNames` setting. A loopback HTTP server serves those actual build files. The emitted classic `pdf.worker.min-*.js` asset must be byte-identical to the installed `pdfjs-dist/build/pdf.worker.min.js` from version 3.11.174. Every served worker response records its real body, SHA-256, size and JavaScript MIME type. The committed, emitted and genuinely served `pdfjs-LICENSE.txt` must also match the installed upstream license byte-for-byte at both bases. There is no intercepted worker fulfillment, CDN substitution or replacement renderer.

The browser's native Worker constructor is observed without replacing its implementation. Scope evaluation starts only for the worker from the current conversion after its successful native PDF.js handshake; failed-load workers are recorded without requesting an execution context. Scope evaluations and supplemental browser-body reads have explicit 10-second deadlines. Missing-worker phase records and partial scenario evidence are written before bounded context cleanup, so an unavailable observer cannot hide the last reached phase. Every accepted conversion must construct a same-origin classic worker, receive PDF.js's successful native transferable-data handshake, and execute in a dedicated worker scope without a document. Native worker event URLs are normalized against the actual page URL, with the raw URL retained, and must still match the constructor and exact deployment asset. The worker URL must preserve the configured deployment base. Fake-worker fallback is forbidden on healthy conversions, repeated conversions, canvas-fault retries and recovery after reload. External HTTP requests and all non-read HTTP requests are blocked and fail acceptance. Only development Vite's same-origin websocket is permitted.

## Conversion coverage

Each of the three engines runs 7 scenarios on each production base and 6 scenarios against the actual Vite development server: 60 scenario results in total.

- Both original three-page scanned and text/vector PDF fixtures convert successfully and then repeat with the same selected file on the same page
- Page 2 returns a null canvas context, an empty `data:,` export, malformed base64, or an encoding exception
- Each canvas fault produces a page-specific error, no download, no success message, no 100% status and no processing of page 3; restoring the canvas boundary then retries on the same page with the same selected file
- Production worker absence is a real HTTP 404 at the emitted worker URL, including fallback requests; no partial download or false completion is allowed
- After restoring the worker resource, the suite deliberately probes the same page and records its actual result. PDF.js 3.11 may retain disabled-worker state and a rejected fake-worker setup promise. A subsequent reload and selection of the identical fixture must recover with a genuine dedicated worker. Reload recovery is explicitly recorded and is never described as in-place recovery
- Every accepted DOCX has exactly three ordered media images with exact byte equality to each real rendered source-page JPEG, page labels, source title, creation date, the existing 600px image width and 2:1 source aspect ratio
- Rendered source pixels establish red/green/blue page order and visible interior text/content, so a solid-color blank render cannot satisfy the content check

Faults remain at the browser canvas boundary. A strict classifier applies only to these two synthetic fixtures: a newly created detached canvas must make its first exact `getContext('2d')` request at the known 720 × 360 viewport, and export exactly once with `toDataURL('image/jpeg', 0.95)` on the same WeakMap identity before the next source canvas starts. Extra, duplicate, reused, out-of-order or interleaved source-sized candidates create independent protocol violations and fail acceptance even if a renderer catches the call's result. An ambiguous PDF.js scratch canvas is never silently selected as a passing source page. Context and export identities, arguments, attempt IDs, page numbers and invocation counts are retained. The null-context fault additionally requires the app's exact page-2 canvas-creation guard message.

Function names in `Error.stack` are not used to classify canvases. The normal production `keepNames` setting remains unchanged. Bounded raw matched/unmatched stack samples are diagnostic only. Unexpected success or a download in a negative case is detected promptly; its DOCX, canvas observations and source images are retained before the case fails. Every failure prints a bounded message with scenario and last phase. PDF.js, React, DOCX packing, FileSaver and downloads remain real.


## Review evidence

Artifacts include actual builds, served worker bytes, each downloaded DOCX, rendered source JPEGs, screenshots, visible page text, exact canvas/page/attempt identities and bounded diagnostic stacks, native worker evidence, network/server ledgers and hashes. `summary.json` records the checked-out commit/tree, workflow SHA and run/attempt, browser/runtime versions, package and lockfile hashes, fixture hashes, deployment bases and every result. `manifest.json` binds each evidence file to that commit/tree and SHA-256. Browser response bodies are cross-checked when exposed by the engine; the real-origin response ledger and captured worker body are always required for production.

Review screenshots and render downloaded DOCX files in a document viewer after downloading the artifact. ZIP/media assertions do not replace visual acceptance of Word or LibreOffice layout. A script that is merely configured, syntax-checked or built has not passed hosted browser acceptance. Focused non-browser orchestration tests run with `node --test tests/browser/observation.test.cjs tests/browser/canvas-boundary.test.cjs`; they exercise missing handshakes, never-ready and closed workers, successful scope observations, current-attempt selection, bounded cleanup, all canvas lifecycle violations, URL normalization and prompt unexpected-success/download recognition. A never-settling mock also demonstrates why the former cumulative `Promise.all` could remain pending; this establishes the harness risk without proving any particular hosted-run cause.

The synthetic PDF fixtures contain no user documents. Each has three 480 × 240 point pages and the red, green and blue page colors (220, 50, 50), (40, 150, 60) and (45, 85, 210). Their unchanged hashes are:

- `scanned-three-pages.pdf`: `3271e2e9dd398e37a12dbfd0d917f850a8d5c7fc59ab9a9c49f7f7e40f2b8661`
- `vector-three-pages.pdf`: `9fab680ced24e8b13a6ac2243c745238be0cfdda33ec0ccf86f47e1c2451d264`

## Existing dependency audit failures

The application lockfile is unchanged from baseline commit `e00efa9092aab809b40e1336a0b1672174ad47f2` (SHA-256 `5d5c0817fb159cc0444d0576c0f7d50009be75b77b96c0836b79a8acd1cfbe11`). The recorded baseline full audit fails with 31 findings: 1 low, 4 moderate, 25 high, and 1 critical. The production-only audit fails with 17 findings: 2 moderate, 14 high, and 1 critical. These are existing, unresolved findings; the rendering configuration mitigation does not make either audit pass.

The build, lint, conversion tests, and isolated browser-tool audit are separate checks. Their passing results must not be described as a passing application dependency audit. The existing full and production audit JSON/exit-code records are retained with the repair assessment. No runtime dependency upgrade or audit-gate suppression is part of this change.
