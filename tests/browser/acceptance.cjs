const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { createRequire } = require('node:module');
const { execFileSync, spawn } = require('node:child_process');
const { chromium } = require('playwright');

const root = path.resolve(__dirname, '../..');
const requireApp = createRequire(path.join(root, 'package.json'));
const JSZip = createRequire(requireApp.resolve('docx'))('jszip');
const output = path.join(__dirname, 'artifacts');
const baseURL = 'http://127.0.0.1:4317';
const workerURL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
const successText = 'Chuyển đổi hoàn tất! Tệp DOCX đã được tải về.';
const buttonName = 'Chuyển đổi sang Word';
const scenarios = [
  { name: 'healthy-scanned', fixture: 'scanned-three-pages.pdf' },
  { name: 'healthy-vector', fixture: 'vector-three-pages.pdf' },
  ...['null-context', 'empty-image', 'malformed-base64', 'image-encoding'].map(fault => ({
    name: `page-2-${fault}`, fixture: 'scanned-three-pages.pdf', fault,
  })),
];
const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
const writeJson = (file, value) => fs.writeFile(file, JSON.stringify(value, null, 2) + '\n');
const git = (...args) => execFileSync('git', args, { cwd: root, encoding: 'utf8' }).trim();

// Narrow faults at the browser boundary, never replacing PDF.js, docx, React,
// or the converter. Vite serves its normal, non-minified source module.
function observeCanvas() {
  const evidence = window.__pdfAcceptance = { fault: null, canvases: 0, exports: [], faults: [] };
  const nativeContext = HTMLCanvasElement.prototype.getContext;
  const nativeExport = HTMLCanvasElement.prototype.toDataURL;
  HTMLCanvasElement.prototype.getContext = function (...args) {
    const caller = new Error().stack.split('\n')[2] || '';
    if (args[0] === '2d' && caller.includes('convertPdfToWord')) {
      this.__pdfSourcePage = ++evidence.canvases;
      if (this.__pdfSourcePage === 2 && evidence.fault === 'null-context') {
        evidence.faults.push({ page: 2, fault: evidence.fault });
        return null;
      }
    }
    return nativeContext.apply(this, args);
  };
  HTMLCanvasElement.prototype.toDataURL = function (...args) {
    const actual = nativeExport.apply(this, args);
    if (!this.__pdfSourcePage) return actual;
    const pixel = nativeContext.call(this, '2d').getImageData(30, 30, 1, 1).data;
    evidence.exports.push({ page: this.__pdfSourcePage, width: this.width, height: this.height,
      sample: Array.from(pixel), actual });
    if (this.__pdfSourcePage === 2 && evidence.fault) {
      evidence.faults.push({ page: 2, fault: evidence.fault });
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

async function assertDocument(bytes, sourceImages, sourceState, fixture, directory) {
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file('word/document.xml').async('string');
  const rels = await zip.file('word/_rels/document.xml.rels').async('string');
  const relationships = new Map([...rels.matchAll(/<Relationship\b[^>]*\/>/g)].map(([tag]) => [
    /\bId="([^"]+)"/.exec(tag)[1], /\bTarget="([^"]+)"/.exec(tag)[1],
  ]));
  const images = [...xml.matchAll(/<a:blip[^>]*\br:embed="([^"]+)"/g)].map(match => match[1]);
  const media = Object.keys(zip.files).filter(name => name.startsWith('word/media/') && !zip.files[name].dir);
  assert.equal(images.length, 3, 'one image per source PDF page');
  assert.equal(media.length, 3, 'no omitted or extra media');
  assert.deepEqual(sourceState.exports.map(item => item.page), [1, 2, 3]);
  const expectedColors = [[220, 50, 50], [40, 150, 60], [45, 85, 210]];
  const orderedMedia = [];
  for (const [index, id] of images.entries()) {
    const target = path.posix.normalize(path.posix.join('word', relationships.get(id)));
    const image = await zip.file(target).async('nodebuffer');
    assert.ok(image.equals(sourceImages[index]), `DOCX image ${index + 1} equals real rendered source page bytes`);
    assert.equal(image.subarray(0, 2).toString('hex'), 'ffd8', 'JPEG signature');
    assert.ok(sourceState.exports[index].sample.slice(0, 3).every((value, channel) =>
      Math.abs(value - expectedColors[index][channel]) <= 6), `source page ${index + 1} has its expected color marker`);
    assert.equal(sourceState.exports[index].width, 720);
    assert.equal(sourceState.exports[index].height, 360);
    orderedMedia.push({ page: index + 1, path: target, sha256: sha256(image), bytes: image.length });
  }
  assert.deepEqual([...xml.matchAll(/<w:t[^>]*>(Trang \d+)<\/w:t>/g)].map(match => match[1]),
    ['Trang 1', 'Trang 2', 'Trang 3']);
  assert.ok(xml.includes(`Tài liệu chuyển đổi từ: ${fixture}`));
  assert.ok(xml.includes('Được tạo vào:'));
  assert.deepEqual([...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"/g)].map(match => match.slice(1)),
    Array.from({ length: 3 }, () => ['5715000', '2857500']), 'retain 600px width and source aspect ratio');
  const report = { docxSha256: sha256(bytes), docxBytes: bytes.length, sourcePages: 3, orderedMedia };
  await writeJson(path.join(directory, 'docx-inspection.json'), report);
  return report;
}

async function arm(page, fault = null) {
  await page.evaluate(nextFault => {
    Object.assign(window.__pdfAcceptance, { fault: nextFault, canvases: 0, exports: [], faults: [] });
  }, fault);
}

async function completeConversion(page, scenario, directory, downloads) {
  const before = downloads.length;
  const pendingDownload = page.waitForEvent('download', { timeout: 45000 });
  // Install the listener before triggering the actual UI download.
  const [download] = await Promise.all([
    pendingDownload,
    page.getByRole('button', { name: buttonName, exact: true }).click(),
  ]);
  const filename = scenario.fixture.replace('.pdf', '_converted.docx');
  assert.equal(download.suggestedFilename(), filename);
  await fs.mkdir(directory, { recursive: true });
  const target = path.join(directory, filename);
  await download.saveAs(target);
  assert.equal(await download.failure(), null);
  await page.getByText(successText, { exact: true }).waitFor();
  await page.waitForFunction(name => [...document.querySelectorAll('button')].some(button =>
    button.textContent === name && !button.disabled), buttonName);
  const recorded = await recordState(page, directory);
  assert.equal(downloads.length, before + 1, 'exactly one download per successful attempt');
  assert.equal(recorded.state.canvases, 3);
  assert.deepEqual(recorded.state.faults, []);
  assert.ok(!recorded.body.includes('Lỗi khi chuyển đổi:'));
  return assertDocument(await fs.readFile(target), recorded.images, recorded.state, scenario.fixture, directory);
}

async function runScenario(browser, scenario) {
  const directory = path.join(output, scenario.name);
  await fs.mkdir(directory, { recursive: true });
  const context = await browser.newContext({ acceptDownloads: true, serviceWorkers: 'block',
    viewport: { width: 1150, height: 950 } });
  const network = [], errors = [], logs = [], workers = [], downloads = [];
  await context.route('**/*', async route => {
    const request = route.request();
    const url = request.url();
    if (url === workerURL && request.method() === 'GET') {
      network.push({ url, action: 'same-version-installed-worker' });
      return route.fulfill({ path: requireApp.resolve('pdfjs-dist/build/pdf.worker.min.js'),
        contentType: 'application/javascript', headers: { 'Access-Control-Allow-Origin': '*' } });
    }
    if (new URL(url).origin === baseURL && ['GET', 'HEAD'].includes(request.method())) return route.continue();
    network.push({ url, method: request.method(), resourceType: request.resourceType(), action: 'blocked-unexpected-request' });
    return route.abort('blockedbyclient');
  });
  await context.routeWebSocket('**/*', socket => {
    if (new URL(socket.url()).origin === baseURL.replace('http:', 'ws:')) {
      socket.connectToServer();
    } else {
      network.push({ url: socket.url(), resourceType: 'websocket', action: 'blocked-unexpected-request' });
      socket.close();
    }
  });
  await context.addInitScript(observeCanvas);
  const page = await context.newPage();
  page.on('pageerror', error => errors.push(String(error)));
  page.on('console', message => logs.push({ type: message.type(), text: message.text() }));
  page.on('worker', worker => workers.push(worker.url()));
  page.on('download', download => downloads.push(download.suggestedFilename()));
  const result = { name: scenario.name, fixture: scenario.fixture, status: 'running' };
  try {
    await page.goto(baseURL, { waitUntil: 'networkidle' });
    await page.locator('input[type=file]').setInputFiles(path.join(__dirname, 'fixtures', scenario.fixture));
    await arm(page, scenario.fault);
    if (scenario.fault) {
      await page.getByRole('button', { name: buttonName, exact: true }).click();
      const errorAlert = page.getByRole('alert');
      await errorAlert.waitFor({ state: 'visible' });
      assert.match(await errorAlert.innerText(), /^Lỗi khi chuyển đổi: Không thể chuyển đổi trang 2:/);
      await page.waitForFunction(name => [...document.querySelectorAll('button')].some(button =>
        button.textContent === name && !button.disabled), buttonName);
      // FileSaver schedules its click; allow that task to settle before retrying.
      await page.waitForTimeout(250);
      const failed = await recordState(page, path.join(directory, 'failure'));
      assert.equal(downloads.length, 0, 'failure must never download a partial DOCX');
      assert.ok(!failed.body.includes(successText), 'failure must never claim success');
      assert.ok(!failed.body.includes('100%'), 'failure must never display 100%');
      assert.equal(failed.state.canvases, 2, 'abort before processing page 3');
      assert.deepEqual(failed.state.faults, [{ page: 2, fault: scenario.fault }], 'fault reached the intended page');
      result.failure = { downloads: 0, success: false, page: 2, fault: scenario.fault };
      await arm(page);
      result.retry = await completeConversion(page, scenario, path.join(directory, 'retry'), downloads);
    } else {
      result.healthy = await completeConversion(page, scenario, path.join(directory, 'healthy'), downloads);
    }
    assert.deepEqual(errors, [], 'no uncaught browser errors');
    assert.ok(workers.length > 0, 'PDF.js must use a real dedicated browser worker');
    assert.ok(network.some(request => request.action === 'same-version-installed-worker'));
    assert.ok(!network.some(request => request.action === 'blocked-unexpected-request'), 'unexpected product network request');
    assert.ok(!logs.some(message => /Setting up fake worker/.test(message.text)), 'no PDF.js fake worker fallback');
    result.status = 'passed';
  } catch (error) {
    result.status = 'failed';
    result.error = error.stack;
    await page.screenshot({ path: path.join(directory, 'failure-diagnostic.png'), fullPage: true }).catch(() => {});
  } finally {
    Object.assign(result, { network, errors, logs, workers, downloads });
    await writeJson(path.join(directory, 'result.json'), result);
    await context.close();
  }
  return result;
}

async function artifactManifest(directory) {
  const items = [];
  async function visit(current) {
    for (const entry of await fs.readdir(current, { withFileTypes: true })) {
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) await visit(full);
      else if (entry.name !== 'manifest.json') {
        const bytes = await fs.readFile(full);
        items.push({ path: path.relative(directory, full), bytes: bytes.length, sha256: sha256(bytes) });
      }
    }
  }
  await visit(directory);
  return items.sort((a, b) => a.path.localeCompare(b.path));
}

async function main() {
  assert.equal(process.env.GITHUB_ACTIONS, 'true', 'Run this acceptance in a hosted GitHub Actions runner');
  assert.equal(process.version, 'v24.19.0');
  assert.equal(execFileSync('npm', ['--version'], { encoding: 'utf8' }).trim(), '11.9.0');
  assert.equal(requireApp('pdfjs-dist/package.json').version, '3.11.174', 'worker and PDF.js versions must match');
  assert.equal(git('status', '--porcelain', '--untracked-files=no'), '', 'tracked checkout must be clean');
  await fs.mkdir(output, { recursive: true });
  const summary = { startedAt: new Date().toISOString(), commit: git('rev-parse', 'HEAD'),
    tree: git('rev-parse', 'HEAD^{tree}'), eventSha: process.env.GITHUB_SHA,
    runId: process.env.GITHUB_RUN_ID, runAttempt: process.env.GITHUB_RUN_ATTEMPT,
    node: process.version, npm: '11.9.0', playwright: require('playwright/package.json').version,
    runtimeLockSha256: sha256(await fs.readFile(path.join(root, 'package-lock.json'))),
    toolingLockSha256: sha256(await fs.readFile(path.join(__dirname, 'package-lock.json'))),
    baselineAudit: 'Known full and production audit failures remain in the baseline evidence; browser acceptance is not a runtime dependency audit pass.',
    fixtures: [], scenarios: [], status: 'running' };
  assert.equal(summary.commit, process.env.EXPECTED_COMMIT, 'acceptance must run against the requested commit');
  for (const fixture of ['scanned-three-pages.pdf', 'vector-three-pages.pdf']) {
    summary.fixtures.push({ path: fixture, sha256: sha256(await fs.readFile(path.join(__dirname, 'fixtures', fixture))) });
  }
  const viteLog = await fs.open(path.join(output, 'vite.log'), 'w');
  const viteCli = path.join(path.dirname(requireApp.resolve('vite/package.json')), 'bin/vite.js');
  const server = spawn(process.execPath, [viteCli, '--host', '127.0.0.1', '--port', '4317', '--strictPort'],
    { cwd: root, stdio: ['ignore', viteLog.fd, viteLog.fd] });
  let serverError;
  server.on('error', error => { serverError = error; });
  let browser;
  try {
    const deadline = Date.now() + 30000;
    while (true) {
      if (serverError) throw serverError;
      if (server.exitCode !== null) throw new Error(`Vite exited with ${server.exitCode}`);
      try { if ((await fetch(baseURL)).ok) break; } catch { /* Wait for Vite startup. */ }
      if (Date.now() > deadline) throw new Error('Vite startup timed out');
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    browser = await chromium.launch({ headless: true });
    summary.chromium = browser.version();
    for (const scenario of scenarios) {
      const result = await runScenario(browser, scenario);
      summary.scenarios.push(result);
      console.log(`${result.status}: ${scenario.name}`);
    }
    summary.status = summary.scenarios.every(result => result.status === 'passed') ? 'passed' : 'failed';
  } catch (error) {
    summary.status = 'failed';
    summary.error = error.stack;
  } finally {
    if (browser) await browser.close();
    if (server.pid && server.exitCode === null && server.signalCode === null) {
      await new Promise(resolve => {
        const timeout = setTimeout(() => server.kill('SIGKILL'), 5000);
        server.once('exit', () => { clearTimeout(timeout); resolve(); });
        server.kill('SIGTERM');
      });
    }
    await viteLog.close();
    summary.finishedAt = new Date().toISOString();
    await writeJson(path.join(output, 'summary.json'), summary);
    await writeJson(path.join(output, 'manifest.json'), { commit: summary.commit, tree: summary.tree,
      artifacts: await artifactManifest(output) });
  }
  assert.equal(summary.status, 'passed', `Hosted acceptance failed; inspect ${output}`);
}

main().catch(error => { console.error(error); process.exitCode = 1; });
