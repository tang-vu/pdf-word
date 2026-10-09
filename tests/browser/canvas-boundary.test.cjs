const assert = require('node:assert/strict');
const test = require('node:test');
const { createCanvasBoundary } = require('./canvas-boundary.cjs');
const { normalizeWorkerUrl, selectAttemptWorker, waitForConversionOutcome } = require('./observation.cjs');

function setup() {
  const evidence = { attempt: 1, fixture: { pageCount: 3,
    pages: Array.from({ length: 3 }, () => ({ widthPt: 480, heightPt: 240 })) },
  canvases: 0, contextCalls: [], exportCalls: 0, unmatched: [], protocolViolations: [] };
  const boundary = createCanvasBoundary(evidence);
  const canvas = (width = 720, height = 360) => {
    const value = { width, height, isConnected: false };
    boundary.register(value);
    return value;
  };
  return { evidence, boundary, canvas };
}
const jpeg = ['image/jpeg', 0.95];

test('fixture protocol proves distinct same-canvas context/export identities in order without stack names', () => {
  const { evidence, boundary, canvas } = setup();
  for (let page = 1; page <= 3; page++) {
    const target = canvas();
    const context = boundary.context(target, ['2d'], 'minifiedCaller@file:1:2');
    const image = boundary.export(target, jpeg, 'unrelated engine stack spelling');
    assert.equal(context.page, page);
    assert.equal(image.page, page);
    assert.equal(image.canvasId, context.canvasId);
    assert.equal(image.invocation, page);
  }
  assert.equal(evidence.canvases, 3);
  assert.equal(evidence.exportCalls, 3);
  assert.equal(new Set(evidence.contextCalls.map(call => call.canvasId)).size, 3);
  assert.deepEqual(evidence.protocolViolations, []);
});

test('duplicate context request is recorded even if its caller ignores the result', () => {
  const { evidence, boundary, canvas } = setup();
  const target = canvas();
  boundary.context(target, ['2d']);
  assert.equal(boundary.context(target, ['2d']), null);
  boundary.export(target, jpeg);
  assert.match(evidence.protocolViolations[0].reason, /duplicate context/);
});

test('same-sized internal scratch canvas interleaved before source export cannot pass', () => {
  const { evidence, boundary, canvas } = setup();
  const source = canvas(), scratch = canvas();
  boundary.context(source, ['2d']);
  assert.equal(boundary.context(scratch, ['2d']), null);
  boundary.export(source, jpeg);
  assert.equal(evidence.canvases, 1);
  assert.match(evidence.protocolViolations[0].reason, /interleaved/);
});

test('same-sized scratch canvas selected first makes the later app canvas ambiguous and fails', () => {
  const { evidence, boundary, canvas } = setup();
  boundary.context(canvas(), ['2d']);
  assert.equal(boundary.context(canvas(), ['2d']), null);
  assert.match(evidence.protocolViolations[0].reason, /interleaved/);
});

test('export before context and extra export are independently recorded violations', () => {
  const { evidence, boundary, canvas } = setup();
  assert.equal(boundary.export(canvas(), jpeg), null);
  const source = canvas();
  boundary.context(source, ['2d']);
  boundary.export(source, jpeg);
  assert.equal(boundary.export(source, jpeg), null);
  assert.equal(evidence.protocolViolations.length, 2);
  assert.match(evidence.protocolViolations[0].reason, /without a current/);
  assert.match(evidence.protocolViolations[1].reason, /duplicate/);
});

test('a fourth candidate after all three pages cannot pass', () => {
  const { evidence, boundary, canvas } = setup();
  for (let page = 1; page <= 3; page++) {
    const source = canvas(); boundary.context(source, ['2d']); boundary.export(source, jpeg);
  }
  assert.equal(boundary.context(canvas(), ['2d']), null);
  assert.match(evidence.protocolViolations[0].reason, /after the final page/);
});

test('wrong dimensions/context arguments never identify a source and source-like exports fail', () => {
  const { evidence, boundary, canvas } = setup();
  assert.equal(boundary.context(canvas(480, 240), ['2d']), null);
  const wrongArguments = canvas();
  assert.equal(boundary.context(wrongArguments, ['2d', {}]), null);
  assert.equal(boundary.export(wrongArguments, jpeg), null);
  assert.equal(evidence.canvases, 0);
  assert.equal(evidence.exportCalls, 0);
  assert.equal(evidence.unmatched.length, 3);
  assert.equal(evidence.protocolViolations.length, 1);
});

test('changing source dimensions or export arguments fails without accepting an image', () => {
  for (const change of ['dimensions', 'arguments']) {
    const { evidence, boundary, canvas } = setup();
    const source = canvas(); boundary.context(source, ['2d']);
    if (change === 'dimensions') source.height++;
    assert.equal(boundary.export(source, change === 'arguments' ? ['image/png'] : jpeg), null);
    assert.equal(evidence.exportCalls, 0);
    assert.match(evidence.protocolViolations[0].reason, /arguments or viewport changed/);
  }
});

test('attempt reset rejects prior canvases and gives new source identities fresh page numbers', () => {
  const { evidence, boundary, canvas } = setup();
  const previous = canvas(); boundary.context(previous, ['2d']); boundary.export(previous, jpeg);
  Object.assign(evidence, { attempt: 2, canvases: 0, contextCalls: [], exportCalls: 0,
    unmatched: [], protocolViolations: [] });
  assert.equal(boundary.context(previous, ['2d']), null);
  assert.equal(boundary.export(previous, jpeg), null);
  assert.equal(evidence.protocolViolations.length, 2);
  const current = canvas();
  const context = boundary.context(current, ['2d']);
  assert.equal(context.page, 1);
  assert.equal(context.createdAttempt, 2);
  assert.notEqual(context.canvasId, 1);
  assert.equal(boundary.export(current, jpeg).invocation, 1);
});

test('unmatched diagnostic stacks are bounded and never establish identity', () => {
  const { evidence, boundary, canvas } = setup();
  for (let i = 0; i < 30; i++) boundary.context(canvas(1, 1), ['2d'], 'convertPdfToWord'.repeat(400));
  assert.equal(evidence.canvases, 0);
  assert.equal(evidence.unmatched.length, 12);
  assert.ok(evidence.unmatched.every(sample => sample.stack.length <= 1800));
});

test('relative and absolute worker event URLs normalize against the actual page', () => {
  const page = 'http://127.0.0.1:4317/pdf-word/';
  const expected = 'http://127.0.0.1:4317/pdf-word/assets/pdf.worker.min-hash.js';
  assert.deepEqual(normalizeWorkerUrl('/pdf-word/assets/pdf.worker.min-hash.js', page),
    { rawUrl: '/pdf-word/assets/pdf.worker.min-hash.js', url: expected });
  assert.equal(normalizeWorkerUrl('assets/pdf.worker.min-hash.js', page).url, expected);
  assert.equal(normalizeWorkerUrl(expected, page).url, expected);
});

test('normalization cannot hide another origin/base or select an earlier same-URL worker', () => {
  const page = 'http://127.0.0.1:4317/pdf-word/';
  const expected = normalizeWorkerUrl('assets/worker.js', page).url;
  const old = { id: 0, ...normalizeWorkerUrl('assets/worker.js', page) };
  for (const raw of ['https://example.com/pdf-word/assets/worker.js', '/assets/worker.js']) {
    const wrong = { id: 1, ...normalizeWorkerUrl(raw, page) };
    assert.throws(() => selectAttemptWorker([old, wrong], 1, expected), /observed 0/);
  }
  const current = { id: 2, ...normalizeWorkerUrl('assets/worker.js', page) };
  assert.equal(selectAttemptWorker([old, current], 1, expected).id, 2);
});

test('negative probes recognize unexpected success/download on first observation', async () => {
  for (const state of [{ success: true }, { downloadCount: 1 }]) {
    let reads = 0;
    const result = await waitForConversionOutcome(async () => { reads++; return state; }, 1000);
    assert.equal(reads, 1);
    assert.equal(result.kind, state.success ? 'success' : 'download');
  }
});

test('negative probe distinguishes the expected error and times out when nothing happens', async () => {
  assert.equal((await waitForConversionOutcome(async () => ({ error: 'expected page 2 guard' }), 50)).kind, 'error');
  await assert.rejects(waitForConversionOutcome(async () => ({}), 20, 2), /no error, success or download/);
});
