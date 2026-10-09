const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const http = require('node:http');
const { createHash } = require('node:crypto');
const { createRequire } = require('node:module');
const { execFileSync, spawn } = require('node:child_process');
const playwright = require('playwright');

const root = path.resolve(__dirname, '../..');
const requireApp = createRequire(path.join(root, 'package.json'));
const JSZip = createRequire(requireApp.resolve('docx'))('jszip');
const output = path.join(__dirname, 'artifacts');
const fixtureDirectory = path.join(__dirname, 'fixtures');
const origin = 'http://127.0.0.1:4317';
const version = '3.11.174';
const pdfPackage = path.dirname(requireApp.resolve('pdfjs-dist/package.json'));
const workerRelative = 'build/pdf.worker.min.js';
const successText = 'Chuyển đổi hoàn tất! Tệp DOCX đã được tải về.';
const buttonName = 'Chuyển đổi sang Word';
const faultNames = ['null-context', 'empty-image', 'malformed-base64', 'image-encoding'];
const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
const writeJson = (file, value) => fs.writeFile(file, JSON.stringify(value, null, 2) + '\n');
const git = (...args) => execFileSync('git', args, { cwd: root, encoding: 'utf8' }).trim();
const pause = milliseconds => new Promise(resolve => setTimeout(resolve, milliseconds));
const mime = filename => ({ '.html': 'text/html', '.js': 'application/javascript',
  '.mjs': 'application/javascript', '.css': 'text/css', '.wasm': 'application/wasm',
  '.ttf': 'font/ttf', '.otf': 'font/otf', '.icc': 'application/vnd.iccprofile',
  '.svg': 'image/svg+xml', '.json': 'application/json', '.txt': 'text/plain' }[path.extname(filename)] || 'application/octet-stream');

// Observe native browser boundaries. Arguments and return values remain untouched
// except for the explicitly armed page-2 fault. keepNames is the production config,
// not a testing-only build override. Record the immediate caller on every page.
function observeBrowser() {
  const evidence = window.__pdfAcceptance = { fault: null, canvases: 0, exports: [],
    faults: [], contextCalls: [], exportCalls: 0, workersConstructed: [], fixture: null, attempt: 0 };
  const NativeWorker = window.Worker;
  window.Worker = new Proxy(NativeWorker, {
    construct(target, args, newTarget) {
      const observation = { url: new URL(String(args[0]), location.href).href,
        type: args[1]?.type || 'classic', argumentCount: args.length,
        attempt: evidence.attempt, messages: [], errors: [] };
      evidence.workersConstructed.push(observation);
      const worker = Reflect.construct(target, args, newTarget);
      worker.addEventListener('message', event => {
        const data = event.data;
        if (data && ['ready', 'test'].includes(data.action)) {
          observation.messages.push({ action: data.action, sourceName: data.sourceName,
            targetName: data.targetName, data: data.data });
        }
      });
      worker.addEventListener('error', event => observation.errors.push(event.message || 'worker load error'));
      return worker;
    },
  });
  const nativeContext = HTMLCanvasElement.prototype.getContext;
  const nativeExport = HTMLCanvasElement.prototype.toDataURL;
  HTMLCanvasElement.prototype.getContext = function observedGetContext(...args) {
    const frames = (new Error().stack || '').split('\n');
    const index = frames.findIndex(frame => frame.includes('observedGetContext'));
    const caller = index >= 0 ? frames[index + 1] || '' : '';
    // Chromium prefixes stacks with Error; Firefox/WebKit do not. Find our
    // named wrapper, then only its immediate caller, never any ancestor.
    if (args[0] === '2d' && caller.includes('convertPdfToWord') && !this.__pdfSourcePage) {
      this.__pdfSourcePage = ++evidence.canvases;
      evidence.contextCalls.push({ page: this.__pdfSourcePage, invocation: evidence.canvases, caller });
      if (this.__pdfSourcePage === 2 && evidence.fault === 'null-context') {
        evidence.faults.push({ page: 2, fault: evidence.fault, invocation: evidence.canvases, boundary: 'getContext' });
        return null;
      }
    }
    return nativeContext.apply(this, args);
  };
  HTMLCanvasElement.prototype.toDataURL = function observedToDataURL(...args) {
    const actual = nativeExport.apply(this, args);
    if (!this.__pdfSourcePage) return actual;
    const frames = (new Error().stack || '').split('\n');
    const index = frames.findIndex(frame => frame.includes('observedToDataURL'));
    const caller = index >= 0 ? frames[index + 1] || '' : '';
    const page = this.__pdfSourcePage;
    const control = evidence.fixture.pages[page - 1];
    const context = nativeContext.call(this, '2d');
    const scaleX = this.width / control.widthPt, scaleY = this.height / control.heightPt;
    const samples = (control.samples || []).map(sample => ({ name: sample.name,
      rgba: Array.from(context.getImageData(Math.floor(sample.xPt * scaleX),
        Math.floor(sample.yPt * scaleY), 1, 1).data) }));
    const regions = (control.regions || []).map(region => {
      const pixels = context.getImageData(Math.floor(region.xPt * scaleX), Math.floor(region.yPt * scaleY),
        Math.max(1, Math.floor(region.widthPt * scaleX)), Math.max(1, Math.floor(region.heightPt * scaleY))).data;
      let white = 0;
      for (let i = 0; i < pixels.length; i += 4) {
        if (pixels[i] > 235 && pixels[i + 1] > 235 && pixels[i + 2] > 235) white++;
      }
      return { name: region.name, whiteFraction: white / (pixels.length / 4) };
    });
    const invocation = ++evidence.exportCalls;
    evidence.exports.push({ page, invocation, caller, width: this.width, height: this.height, samples, regions, actual });
    if (page === 2 && evidence.fault) {
      evidence.faults.push({ page: 2, fault: evidence.fault, invocation, boundary: 'toDataURL' });
      if (evidence.fault === 'empty-image') return 'data:,';
      if (evidence.fault === 'malformed-base64') return 'data:image/jpeg;base64,%';
      if (evidence.fault === 'image-encoding') throw new Error('Controlled page 2 image encoding failure');
    }
    return actual;
  };
}

async function recordState(page, directory) {
  await fs.mkdir(directory, { recursive: true });
  const state = await page.evaluate(() => window.__pdfAcceptance);
  const images = [];
  for (const item of state.exports) {
    const bytes = Buffer.from(item.actual.split(',')[1], 'base64');
    const filename = `rendered-source-page-${item.page}.jpg`;
    await fs.writeFile(path.join(directory, filename), bytes);
    images.push(bytes);
    delete item.actual;
    item.filename = filename;
    item.sha256 = sha256(bytes);
  }
  const body = await page.locator('body').innerText();
  await fs.writeFile(path.join(directory, 'page.txt'), body);
  await page.screenshot({ path: path.join(directory, 'conversion.png'), fullPage: true });
  await writeJson(path.join(directory, 'canvas.json'), state);
  return { state, images, body };
}

function assertPixels(sourceState, fixture) {
  for (const [index, actual] of sourceState.exports.entries()) {
    const expected = fixture.pages[index];
    assert.equal(actual.width, expected.widthPt * 1.5, 'normal source viewport width');
    assert.equal(actual.height, expected.heightPt * 1.5, 'normal source viewport height');
    for (const [sampleIndex, sample] of (expected.samples || []).entries()) {
      const observed = actual.samples[sampleIndex];
      assert.equal(observed.name, sample.name);
      assert.ok(observed.rgba.slice(0, 3).every((value, channel) =>
        Math.abs(value - sample.rgb[channel]) <= (sample.tolerance ?? 6)),
      `${fixture.file} page ${index + 1} ${sample.name}: expected ${sample.rgb}, got ${observed.rgba}`);
    }
    for (const [regionIndex, region] of (expected.regions || []).entries()) {
      const fraction = actual.regions[regionIndex].whiteFraction;
      assert.ok(fraction >= region.minFraction, `${fixture.file} ${region.name}: visible source content ${fraction} < ${region.minFraction}`);
      if (region.maxFraction !== undefined) assert.ok(fraction <= region.maxFraction, 'source region upper bound');
    }
  }
}

async function assertDocument(bytes, recorded, fixture, directory, checkPixels = true) {
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file('word/document.xml').async('string');
  const rels = await zip.file('word/_rels/document.xml.rels').async('string');
  const relationships = new Map([...rels.matchAll(/<Relationship\b[^>]*\/>/g)].map(([tag]) => [
    /\bId="([^"]+)"/.exec(tag)[1], /\bTarget="([^"]+)"/.exec(tag)[1],
  ]));
  const images = [...xml.matchAll(/<a:blip[^>]*\br:embed="([^"]+)"/g)].map(match => match[1]);
  const media = Object.keys(zip.files).filter(name => name.startsWith('word/media/') && !zip.files[name].dir);
  assert.equal(images.length, fixture.pageCount, 'one image per source PDF page');
  assert.equal(media.length, fixture.pageCount, 'no omitted or extra media');
  assert.deepEqual(recorded.state.exports.map(item => item.page), fixture.pages.map((_, index) => index + 1));
  const orderedMedia = [];
  for (const [index, id] of images.entries()) {
    const target = path.posix.normalize(path.posix.join('word', relationships.get(id)));
    const image = await zip.file(target).async('nodebuffer');
    assert.ok(image.equals(recorded.images[index]), `DOCX page ${index + 1} equals real rendered source bytes in order`);
    assert.equal(image.subarray(0, 2).toString('hex'), 'ffd8', 'JPEG signature');
    orderedMedia.push({ page: index + 1, path: target, sha256: sha256(image), bytes: image.length });
  }
  assert.deepEqual([...xml.matchAll(/<w:t[^>]*>(Trang \d+)<\/w:t>/g)].map(match => match[1]),
    fixture.pages.map((_, index) => `Trang ${index + 1}`));
  assert.ok(xml.includes(`Tài liệu chuyển đổi từ: ${fixture.file}`));
  assert.ok(xml.includes('Được tạo vào:'));
  assert.deepEqual([...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"/g)].map(match => match.slice(1)),
    fixture.pages.map(page => ['5715000', String(Math.floor(600 * page.heightPt / page.widthPt) * 9525)]),
    'retain 600px width and source aspect ratio');
  const report = { docxSha256: sha256(bytes), docxBytes: bytes.length, sourcePages: fixture.pageCount,
    orderedMedia, pixelChecks: checkPixels ? 'required' : 'recorded for failure diagnosis' };
  await writeJson(path.join(directory, 'docx-inspection.json'), report);
  if (checkPixels) assertPixels(recorded.state, fixture);
  return report;
}

async function arm(page, fixture, fault = null) {
  await page.evaluate(({ control, nextFault }) => {
    Object.assign(window.__pdfAcceptance, { fault: nextFault, fixture: control, attempt: window.__pdfAcceptance.attempt + 1, canvases: 0,
      exports: [], faults: [], contextCalls: [], exportCalls: 0 });
  }, { control: fixture, nextFault: fault });
}

async function waitUntilIdle(page) {
  await page.waitForFunction(name => [...document.querySelectorAll('button')].some(button =>
    button.textContent === name && !button.disabled), buttonName, { timeout: 45000 });
}

async function saveDownload(page, download, fixture, directory, checkPixels = true) {
  const filename = fixture.file.replace('.pdf', '_converted.docx');
  assert.equal(download.suggestedFilename(), filename);
  await fs.mkdir(directory, { recursive: true });
  const target = path.join(directory, filename);
  await download.saveAs(target);
  assert.equal(await download.failure(), null);
  await page.getByText(successText, { exact: true }).waitFor();
  await waitUntilIdle(page);
  const recorded = await recordState(page, directory);
  assert.equal(recorded.state.canvases, fixture.pageCount, 'observed every app page despite production minification');
  assert.deepEqual(recorded.state.contextCalls.map(call => call.invocation), fixture.pages.map((_, index) => index + 1));
  assert.equal(recorded.state.exportCalls, fixture.pageCount);
  assert.ok(recorded.state.contextCalls.every(call => call.caller.includes('convertPdfToWord')));
  assert.ok(recorded.state.exports.every(call => call.caller.includes('convertPdfToWord')));
  assert.deepEqual(recorded.state.faults, []);
  assert.ok(!recorded.body.includes('Lỗi khi chuyển đổi:'));
  return assertDocument(await fs.readFile(target), recorded, fixture, directory, checkPixels);
}

async function completeConversion(page, fixture, directory, downloads) {
  const before = downloads.length;
  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: 45000 }),
    page.getByRole('button', { name: buttonName, exact: true }).click(),
  ]);
  const report = await saveDownload(page, download, fixture, directory);
  assert.equal(downloads.length, before + 1, 'exactly one genuine browser download per successful attempt');
  return report;
}

async function captureFailure(page, directory, downloads, sourcePage = null) {
  const alert = page.getByRole('alert');
  await alert.waitFor({ state: 'visible', timeout: 45000 });
  assert.match(await alert.innerText(), sourcePage ? new RegExp(`^Lỗi khi chuyển đổi: Không thể chuyển đổi trang ${sourcePage}:`) : /^Lỗi khi chuyển đổi:/);
  await waitUntilIdle(page);
  // FileSaver queues its click; settle that task before claiming no download.
  await page.waitForTimeout(250);
  const recorded = await recordState(page, directory);
  assert.equal(downloads.length, 0, 'failure must never download a partial DOCX');
  assert.ok(!recorded.body.includes(successText), 'failure must never claim success');
  assert.ok(!recorded.body.includes('100%'), 'failure must never display 100%');
  return recorded;
}

async function enumerate(directory) {
  const files = [];
  async function visit(current) {
    for (const entry of await fs.readdir(current, { withFileTypes: true })) {
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) await visit(full); else files.push(path.relative(directory, full).split(path.sep).join('/'));
    }
  }
  await visit(directory);
  return files.sort();
}

async function packageInventory(buildDirectory, base) {
  const workerBytes = await fs.readFile(path.join(pdfPackage, workerRelative));
  const workerFiles = (await enumerate(path.join(buildDirectory, 'assets'))).filter(file => /^pdf\.worker\.min-.*\.js$/.test(file));
  assert.equal(workerFiles.length, 1, 'one Vite-emitted classic worker asset');
  const workerPath = 'assets/' + workerFiles[0];
  assert.deepEqual(await fs.readFile(path.join(buildDirectory, workerPath)), workerBytes,
    'production worker emitted unchanged from the installed package');
  const licenseBytes = await fs.readFile(path.join(pdfPackage, 'LICENSE'));
  assert.deepEqual(await fs.readFile(path.join(root, 'public/pdfjs-LICENSE.txt')), licenseBytes, 'committed upstream license unchanged');
  assert.deepEqual(await fs.readFile(path.join(buildDirectory, 'pdfjs-LICENSE.txt')), licenseBytes, 'emitted upstream license unchanged');
  return [
    { url: origin + base + workerPath, packagePath: workerRelative,
      sha256: sha256(workerBytes), bytes: workerBytes.length, mime: 'application/javascript' },
    { url: origin + base + 'pdfjs-LICENSE.txt', packagePath: 'LICENSE',
      sha256: sha256(licenseBytes), bytes: licenseBytes.length, mime: 'text/plain' },
  ];
}

// This origin serves the normal production build. Missing resources are real
// HTTP 404 responses from this server, never fulfilled browser interception.
async function staticServer(buildDirectory, base, assets, evidenceDirectory) {
  const ledger = [], missing = new Set(), expected = new Map(assets.map(asset => [new URL(asset.url).pathname, asset]));
  await fs.mkdir(path.join(evidenceDirectory, 'served-resources'), { recursive: true });
  const server = http.createServer(async (request, response) => {
    const pathname = new URL(request.url, origin).pathname;
    const entry = { url: origin + request.url, method: request.method, status: 0 };
    ledger.push(entry);
    try {
      if (!['GET', 'HEAD'].includes(request.method)) { entry.status = 405; response.writeHead(405).end(); return; }
      if (!pathname.startsWith(base) || missing.has(pathname)) { entry.status = 404; response.writeHead(404).end('Missing acceptance resource'); return; }
      const relative = decodeURIComponent(pathname.slice(base.length)) || 'index.html';
      const file = path.resolve(buildDirectory, relative);
      if (!file.startsWith(buildDirectory + path.sep)) { entry.status = 403; response.writeHead(403).end(); return; }
      const bytes = await fs.readFile(file);
      entry.status = 200; entry.mime = mime(file); entry.bytes = bytes.length; entry.sha256 = sha256(bytes);
      const asset = expected.get(pathname);
      if (asset) {
        entry.packagePath = asset.packagePath;
        entry.capture = 'actual-origin-response-body';
        entry.bodyFile = 'served-resources/' + asset.packagePath.replaceAll('/', '__');
        await fs.writeFile(path.join(evidenceDirectory, entry.bodyFile), bytes);
      }
      response.writeHead(200, { 'Content-Type': entry.mime, 'Content-Length': bytes.length, 'Cache-Control': 'no-store' });
      response.end(request.method === 'HEAD' ? undefined : bytes);
    } catch (error) {
      entry.status = error.code === 'ENOENT' ? 404 : 500;
      entry.error = String(error);
      response.writeHead(entry.status).end();
    }
  });
  await new Promise((resolve, reject) => { server.once('error', reject); server.listen(4317, '127.0.0.1', resolve); });
  return { ledger, missing, close: () => new Promise(resolve => server.close(resolve)) };
}

function assertResourceResponses(result, deployment, serverEntries) {
  if (deployment.mode !== 'production') return;
  const served = serverEntries.filter(item => item.packagePath && item.status === 200);
  assert.ok(served.some(item => item.packagePath === workerRelative), 'genuine emitted worker was served');
  for (const entry of served) {
    const expected = deployment.assets.find(asset => asset.packagePath === entry.packagePath);
    assert.equal(entry.sha256, expected.sha256, 'actual worker response matches installed package bytes');
    assert.equal(entry.mime, expected.mime, 'actual worker response MIME');
    assert.ok(new URL(entry.url).pathname.startsWith(deployment.base), 'worker respects configured base');
  }
  result.verifiedResources = [...new Set(served.map(item => item.packagePath))];
}

async function verifyWorkerAttempt(page, deployment, workers, pending, logs, logStart) {
  await Promise.all(pending);
  const state = await page.evaluate(() => window.__pdfAcceptance);
  const constructors = state.workersConstructed.filter(worker => worker.attempt === state.attempt);
  assert.equal(constructors.length, 1, 'each conversion constructs its own native worker');
  const observed = constructors[0];
  assert.equal(observed.type, 'classic', 'PDF.js 3.11 uses its native classic worker');
  assert.equal(new URL(observed.url).origin, origin, 'worker stays on the app origin');
  assert.ok(new URL(observed.url).pathname.startsWith(deployment.base), 'worker preserves deployment base');
  if (deployment.mode === 'production') assert.equal(observed.url, deployment.assets[0].url);
  assert.deepEqual(observed.errors, [], 'no native worker errors on accepted conversions');
  assert.ok(observed.messages.some(message => message.action === 'test' && message.data === true &&
    message.sourceName === 'worker' && message.targetName === 'main'), 'real PDF.js worker passed its native transfer handshake');
  const execution = workers.filter(worker => worker.url === observed.url).at(-1);
  assert.ok(execution, 'native constructor produced a Playwright dedicated-worker event');
  assert.equal(execution.scope?.dedicated, true, 'executed inside DedicatedWorkerGlobalScope');
  assert.equal(execution.scope?.documentType, 'undefined', 'worker execution is outside the document');
  assert.equal(execution.scope?.origin, origin);
  assert.ok(!logs.slice(logStart).some(message => /fake worker/i.test(message.text)), 'no fake-worker fallback in accepted conversion');
  return { constructor: observed, execution };
}

async function runScenario(browser, engine, scenario, fixture, deployment) {
  const directory = path.join(output, deployment.name, engine, scenario.name);
  await fs.mkdir(directory, { recursive: true });
  const context = await browser.newContext({ acceptDownloads: true, serviceWorkers: 'block', viewport: { width: 1150, height: 950 } });
  const network = [], errors = [], logs = [], workers = [], downloads = [], responses = [], pending = [];
  const ledgerStart = deployment.server?.ledger.length || 0;
  await context.route('**/*', route => {
    const request = route.request(), url = request.url();
    const allowed = new URL(url).origin === origin && ['GET', 'HEAD'].includes(request.method());
    network.push({ url, method: request.method(), resourceType: request.resourceType(), action: allowed ? 'real-origin-request' : 'blocked-unexpected-request' });
    return allowed ? route.continue() : route.abort('blockedbyclient');
  });
  await context.routeWebSocket('**/*', socket => {
    if (deployment.mode === 'development' && new URL(socket.url()).origin === origin.replace('http:', 'ws:')) socket.connectToServer();
    else { network.push({ url: socket.url(), action: 'blocked-unexpected-request', resourceType: 'websocket' }); socket.close(); }
  });
  context.on('response', response => {
    const asset = deployment.assets?.find(item => item.url === response.url());
    if (!asset || !response.ok()) return;
    pending.push((async () => {
      const evidence = { url: response.url(), status: response.status(), packagePath: asset.packagePath,
        mime: response.headers()['content-type'], capture: 'browser-response-body' };
      responses.push(evidence);
      try {
        let bytes;
        try { bytes = await response.body(); }
        catch (error) { evidence.captureUnavailable = String(error); return; }
        evidence.bytes = bytes.length; evidence.sha256 = sha256(bytes);
        assert.equal(evidence.sha256, asset.sha256, 'browser-observed worker response bytes');
        assert.equal(evidence.mime.split(';')[0], asset.mime, 'browser-observed worker response MIME');
      } catch (error) { evidence.error = String(error); }
    })());
  });
  await context.addInitScript(observeBrowser);
  const page = await context.newPage();
  page.on('pageerror', error => errors.push(String(error)));
  page.on('console', message => logs.push({ type: message.type(), text: message.text() }));
  page.on('worker', worker => {
    const evidence = { url: worker.url() };
    workers.push(evidence);
    pending.push(worker.evaluate(() => ({
      dedicated: typeof DedicatedWorkerGlobalScope !== 'undefined' && self instanceof DedicatedWorkerGlobalScope,
      documentType: typeof document, origin: location.origin,
    })).then(scope => { evidence.scope = scope; }, error => { evidence.error = String(error); }));
  });
  page.on('download', download => downloads.push(download));
  const result = { name: scenario.name, fixture: fixture.file, engine, buildBase: deployment.base, mode: deployment.mode, status: 'running', acceptedWorkers: [] };
  const accept = async label => {
    const logStart = logs.length;
    const report = await completeConversion(page, fixture, path.join(directory, label), downloads);
    const worker = await verifyWorkerAttempt(page, deployment, workers, pending, logs, logStart);
    result.acceptedWorkers.push({ phase: label, ...worker });
    return report;
  };
  try {
    await page.goto(origin + deployment.base, { waitUntil: 'networkidle' });
    await page.locator('input[type=file]').setInputFiles(path.join(fixtureDirectory, fixture.file));
    await arm(page, fixture, scenario.fault);
    if (scenario.missingWorker) {
      const workerPath = new URL(deployment.assets[0].url).pathname;
      deployment.server.missing.add(workerPath);
      await page.getByRole('button', { name: buttonName, exact: true }).click();
      const failed = await captureFailure(page, path.join(directory, 'missing-worker'), downloads);
      assert.equal(failed.state.canvases, 0, 'missing worker fails before rendering source pages');
      assert.ok(deployment.server.ledger.slice(ledgerStart).some(entry =>
        new URL(entry.url).pathname === workerPath && entry.status === 404), 'worker was genuinely unavailable at the origin');
      result.failure = { missingWorker: workerPath, downloads: 0, success: false };
      // PDF.js 3.11 can latch both worker-disabled state and a rejected fallback
      // promise. Observe restoration on the same page before recovering by reload.
      deployment.server.missing.delete(workerPath);
      await arm(page, fixture);
      const restoreLogStart = logs.length;
      await page.getByRole('button', { name: buttonName, exact: true }).click();
      await page.waitForFunction(text => document.querySelector('[role=alert]') || document.body.innerText.includes(text), successText, { timeout: 45000 });
      await waitUntilIdle(page);
      await page.waitForTimeout(250);
      if (await page.getByRole('alert').isVisible()) {
        await captureFailure(page, path.join(directory, 'restored-same-page'), downloads);
        result.samePageRestoration = { outcome: 'error-after-resource-restored', reloadRequired: true };
      } else {
        assert.equal(downloads.length, 1, 'same-page restoration produced exactly one download');
        const document = await saveDownload(page, downloads[0], fixture, path.join(directory, 'restored-same-page'));
        try {
          const worker = await verifyWorkerAttempt(page, deployment, workers, pending, logs, restoreLogStart);
          result.samePageRestoration = { outcome: 'dedicated-worker-recovery', reloadRequired: false, document, worker };
        } catch (error) {
          result.samePageRestoration = { outcome: 'download-without-dedicated-worker-proof', reloadRequired: true, document, workerError: String(error) };
        }
      }
      // A fresh document is an explicit recovery boundary, never an in-place retry claim.
      const recoveryErrorStart = errors.length;
      await page.reload({ waitUntil: 'networkidle' });
      await page.locator('input[type=file]').setInputFiles(path.join(fixtureDirectory, fixture.file));
      await arm(page, fixture);
      result.reloadRecovery = await accept('restored-after-reload');
      assert.deepEqual(errors.slice(recoveryErrorStart), [], 'no uncaught errors after reload recovery');
      result.recoveryBoundary = 'page reload and reselect the identical fixture';
    } else if (scenario.fault) {
      const logStart = logs.length;
      await page.getByRole('button', { name: buttonName, exact: true }).click();
      const failed = await captureFailure(page, path.join(directory, 'failure'), downloads, 2);
      assert.equal(failed.state.canvases, 2, 'abort before processing page 3');
      assert.deepEqual(failed.state.contextCalls.map(call => call.invocation), [1, 2], 'exact source-page context invocation');
      assert.ok(failed.state.contextCalls.every(call => call.caller.includes('convertPdfToWord')), 'immediate app caller observed even after minification');
      assert.equal(failed.state.exportCalls, scenario.fault === 'null-context' ? 1 : 2, 'exact source-page export invocation');
      assert.ok(failed.state.exports.every(call => call.caller.includes('convertPdfToWord')));
      assert.deepEqual(failed.state.faults, [{ page: 2, fault: scenario.fault, invocation: 2,
        boundary: scenario.fault === 'null-context' ? 'getContext' : 'toDataURL' }], 'one fault on source page 2');
      result.failure = { downloads: 0, success: false, page: 2, fault: scenario.fault, invocation: 2,
        worker: await verifyWorkerAttempt(page, deployment, workers, pending, logs, logStart) };
      await arm(page, fixture);
      result.retry = await accept('retry-same-page-same-file');
    } else {
      result.healthy = await accept('healthy');
      await arm(page, fixture);
      result.repeat = await accept('repeat-same-page-same-file');
    }
    await Promise.all(pending);
    if (!scenario.missingWorker) assert.deepEqual(errors, [], 'no uncaught browser errors');
    assert.ok(!network.some(request => request.action === 'blocked-unexpected-request'), 'no unexpected external or non-read product requests');
    assert.ok(!responses.some(response => response.error), 'browser-observed resource bodies and MIME match');
    assertResourceResponses(result, deployment, deployment.server?.ledger.slice(ledgerStart) || []);
    result.status = 'passed';
  } catch (error) {
    result.status = 'failed'; result.error = error.stack;
    await page.screenshot({ path: path.join(directory, 'failure-diagnostic.png'), fullPage: true }).catch(() => {});
  } finally {
    deployment.server?.missing.clear();
    await Promise.all(pending);
    Object.assign(result, { network, errors, logs, workers, responses,
      downloads: downloads.map(download => download.suggestedFilename()),
      serverRequests: deployment.server?.ledger.slice(ledgerStart) || [] });
    await writeJson(path.join(directory, 'result.json'), result);
    await context.close();
  }
  return result;
}

async function runProcess(command, args, logFile) {
  const log = await fs.open(logFile, 'w');
  try {
    await new Promise((resolve, reject) => {
      const child = spawn(command, args, { cwd: root, stdio: ['ignore', log.fd, log.fd] });
      child.once('error', reject);
      child.once('exit', code => code === 0 ? resolve() : reject(new Error(`${command} ${args.join(' ')} exited ${code}; inspect ${logFile}`)));
    });
  } finally { await log.close(); }
}

async function developmentServer() {
  const log = await fs.open(path.join(output, 'development-vite.log'), 'w');
  const viteCli = path.join(path.dirname(requireApp.resolve('vite/package.json')), 'bin/vite.js');
  const server = spawn(process.execPath, [viteCli, '--host', '127.0.0.1', '--port', '4317', '--strictPort'],
    { cwd: root, stdio: ['ignore', log.fd, log.fd] });
  let spawnError;
  server.once('error', error => { spawnError = error; });
  try {
    const deadline = Date.now() + 30000;
    while (true) {
      if (spawnError) throw spawnError;
      if (server.exitCode !== null) throw new Error(`Vite exited with ${server.exitCode}`);
      try { if ((await fetch(origin)).ok) break; } catch { /* Wait for Vite startup. */ }
      if (Date.now() > deadline) throw new Error('Vite startup timed out');
      await pause(100);
    }
  } catch (error) { server.kill(); await log.close(); throw error; }
  return { close: async () => {
    if (server.pid && server.exitCode === null && server.signalCode === null) {
      await new Promise(resolve => {
        const timeout = setTimeout(() => server.kill('SIGKILL'), 5000);
        server.once('exit', () => { clearTimeout(timeout); resolve(); });
        server.kill('SIGTERM');
      });
    }
    await log.close();
  } };
}

async function artifactManifest(directory) {
  const items = [];
  for (const file of await enumerate(directory)) {
    if (file === 'manifest.json') continue;
    const bytes = await fs.readFile(path.join(directory, file));
    items.push({ path: file, bytes: bytes.length, sha256: sha256(bytes) });
  }
  return items;
}

async function main() {
  assert.equal(process.env.GITHUB_ACTIONS, 'true', 'Run only on a genuine hosted GitHub Actions runner; never spoof this flag');
  assert.equal(process.env.RUNNER_ENVIRONMENT, 'github-hosted', 'local and self-hosted browsers are outside this acceptance scope');
  assert.equal(process.version, 'v24.19.0');
  assert.equal(execFileSync('npm', ['--version'], { encoding: 'utf8' }).trim(), '11.9.0');
  assert.equal(requireApp('pdfjs-dist/package.json').version, version);
  assert.equal(require('playwright/package.json').version, '1.64.0');
  assert.equal(git('status', '--porcelain', '--untracked-files=no'), '', 'tracked checkout must be clean');
  await fs.mkdir(output, { recursive: true });
  const summary = { startedAt: new Date().toISOString(), commit: git('rev-parse', 'HEAD'), tree: git('rev-parse', 'HEAD^{tree}'),
    eventSha: process.env.GITHUB_SHA, runId: process.env.GITHUB_RUN_ID, runAttempt: process.env.GITHUB_RUN_ATTEMPT,
    node: process.version, npm: '11.9.0', playwright: '1.64.0', pdfjs: version,
    packageSha256: sha256(await fs.readFile(path.join(root, 'package.json'))),
    runtimeLockSha256: sha256(await fs.readFile(path.join(root, 'package-lock.json'))),
    pdfjsPackageSha256: sha256(await fs.readFile(path.join(pdfPackage, 'package.json'))),
    toolingLockSha256: sha256(await fs.readFile(path.join(__dirname, 'package-lock.json'))),
    auditScope: 'Unchanged app audits retain 31 full and 17 production findings in the baseline assessment. Browser acceptance and the enforced isolated-tool audit do not imply an app audit pass.',
    browserBoundary: 'Current hosted Playwright Chromium, Firefox and WebKit engines; not proof of Safari, iOS or historical minimum versions.',
    fixtures: [], deployments: [], engines: {}, scenarios: [], status: 'running' };
  assert.equal(summary.commit, process.env.EXPECTED_COMMIT, 'acceptance must run against the requested exact commit');
  const fixtures = ['scanned-three-pages.pdf', 'vector-three-pages.pdf'].map(file => ({ file, pageCount: 3,
    pages: [[220, 50, 50], [40, 150, 60], [45, 85, 210]].map((rgb, index) => ({ widthPt: 480, heightPt: 240,
      samples: [{ name: `page-${index + 1}-color-marker`, xPt: 20, yPt: 20, rgb, tolerance: 6 }],
      // Interior excludes the white border: require rendered text/content, not
      // just a correctly colored blank page. Both original fixtures use white text.
      regions: [{ name: 'interior-visible-content', xPt: 25, yPt: 35, widthPt: 400, heightPt: 160,
        minFraction: 0.015, maxFraction: 0.30 }],
    })) }));
  for (const fixture of fixtures) {
    summary.fixtures.push({ ...fixture, sha256: sha256(await fs.readFile(path.join(fixtureDirectory, fixture.file))) });
  }
  const controls = fixtures.map(fixture => ({ name: fixture.file.replace('.pdf', ''), fixture: fixture.file }));
  const faults = faultNames.map(fault => ({ name: `page-2-${fault}`, fixture: fixtures[0].file, fault }));
  const missingWorker = { name: 'missing-local-worker', fixture: fixtures[0].file, missingWorker: true };
  try {
    for (const base of ['/', '/pdf-word/']) {
      const name = base === '/' ? 'production-root' : 'production-nested';
      const buildDirectory = path.join(output, 'builds', name);
      // Same npm build script and configured minifier as normal release output.
      await runProcess('npm', ['run', 'build', '--', '--base', base, '--outDir', buildDirectory], path.join(output, `${name}-build.log`));
      const assets = await packageInventory(buildDirectory, base);
      await writeJson(path.join(output, `${name}-installed-assets.json`), assets);
      summary.deployments.push({ name, base, buildDirectory: path.relative(output, buildDirectory), assets });
      const deployment = { name, base, mode: 'production', assets };
      deployment.server = await staticServer(buildDirectory, base, assets, path.join(output, name));
      try {
        const license = assets.find(asset => asset.packagePath === 'LICENSE');
        const response = await fetch(license.url);
        assert.equal(response.status, 200, 'upstream license is served at this deployment base');
        assert.equal(response.headers.get('content-type'), license.mime);
        assert.equal(sha256(Buffer.from(await response.arrayBuffer())), license.sha256, 'served license matches installed package');
        summary.deployments.at(-1).licenseVerified = true;
        for (const engine of ['chromium', 'firefox', 'webkit']) {
          const browser = await playwright[engine].launch({ headless: true });
          summary.engines[engine] = browser.version();
          try {
            for (const scenario of [...controls, ...faults, missingWorker]) {
              const result = await runScenario(browser, engine, scenario, fixtures.find(item => item.file === scenario.fixture), deployment);
              summary.scenarios.push(result);
              console.log(`${result.status}: ${name}/${engine}/${scenario.name}`);
              await writeJson(path.join(output, 'summary.json'), summary);
            }
          } finally { await browser.close(); }
        }
      } finally {
        await writeJson(path.join(output, name, 'server-requests.json'), deployment.server.ledger);
        await deployment.server.close();
      }
    }
    const server = await developmentServer();
    try {
      for (const engine of ['chromium', 'firefox', 'webkit']) {
        const browser = await playwright[engine].launch({ headless: true });
        try {
          for (const scenario of [...controls, ...faults]) {
            const result = await runScenario(browser, engine, scenario, fixtures.find(item => item.file === scenario.fixture),
              { name: 'development', mode: 'development', base: '/' });
            summary.scenarios.push(result);
            console.log(`${result.status}: development/${engine}/${scenario.name}`);
          }
        } finally { await browser.close(); }
      }
    } finally { await server.close(); }
    summary.status = summary.scenarios.every(result => result.status === 'passed') ? 'passed' : 'failed';
  } catch (error) { summary.status = 'failed'; summary.error = error.stack; }
  finally {
    summary.finishedAt = new Date().toISOString();
    await writeJson(path.join(output, 'summary.json'), summary);
    await writeJson(path.join(output, 'manifest.json'), { commit: summary.commit, tree: summary.tree, artifacts: await artifactManifest(output) });
  }
  assert.equal(summary.scenarios.length, 60, 'all three engines completed production root, nested and development suites');
  assert.equal(summary.status, 'passed', `Hosted acceptance failed; inspect ${output}`);
}

main().catch(error => { console.error(error); process.exitCode = 1; });
