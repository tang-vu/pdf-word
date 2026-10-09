const assert = require('node:assert/strict');
const test = require('node:test');
const { withDeadline, selectAttemptWorker, observeAcceptedWorker } = require('./observation.cjs');

const accepted = { messages: [{ action: 'test', data: true, sourceName: 'worker', targetName: 'main' }] };
const absent = { messages: [] };
const scope = { dedicated: true, documentType: 'undefined', origin: 'http://127.0.0.1:4317' };
const never = () => new Promise(() => {});

test('missing handshake never evaluates a failed-load worker', async () => {
  let evaluations = 0;
  const result = await observeAcceptedWorker({ evaluate() { evaluations++; return never(); } }, absent, 20);
  assert.equal(result.status, 'not-requested');
  assert.equal(evaluations, 0);
  assert.equal(result.scope, undefined);
});

test('claimed-ready worker with no execution context is bounded and cannot pass', async () => {
  const result = await observeAcceptedWorker({ evaluate: never }, accepted, 20);
  assert.equal(result.status, 'timed-out');
  assert.match(result.error, /did not settle within 20ms/);
  assert.equal(result.scope, undefined);
});

test('closed worker rejection is retained and settles', async () => {
  const result = await observeAcceptedWorker({ evaluate: async () => { throw new Error('Worker closed'); } }, accepted, 20);
  assert.equal(result.status, 'rejected');
  assert.match(result.error, /Worker closed/);
  assert.equal(result.scope, undefined);
});

test('healthy accepted worker preserves the actual scope observation', async () => {
  let evaluations = 0;
  const result = await observeAcceptedWorker({ evaluate: async () => { evaluations++; return scope; } }, accepted, 20);
  assert.equal(evaluations, 1);
  assert.deepEqual(result, { status: 'fulfilled', scope });
});

test('old failed worker cannot block the next accepted attempt or cleanup', async () => {
  let failedEvaluations = 0, closed = false;
  const failedHandle = { evaluate() { failedEvaluations++; return never(); } };
  const healthyHandle = { evaluate: async () => scope };
  const workers = [{ id: 0, url: '/worker.js' }];
  const handles = new Map([[0, failedHandle], [1, healthyHandle]]);
  const failed = await observeAcceptedWorker(failedHandle, absent, 20);
  const nextAttemptStart = workers.length;
  workers.push({ id: 1, url: '/worker.js' });
  const selected = selectAttemptWorker(workers, nextAttemptStart, '/worker.js');
  assert.equal(selected.id, 1);
  const healthy = await observeAcceptedWorker(handles.get(selected.id), accepted, 20);
  assert.throws(() => selectAttemptWorker(workers, workers.length, '/worker.js'), /observed 0/);
  await withDeadline(Promise.resolve().then(() => { closed = true; }), 'Context cleanup', 20);
  assert.equal(failedEvaluations, 0);
  assert.equal(failed.status, 'not-requested');
  assert.equal(healthy.status, 'fulfilled');
  assert.equal(closed, true);
});

test('former cumulative Promise.all remains pending for a never-ready mock', async () => {
  const oldPending = [never(), Promise.resolve(scope)];
  await assert.rejects(withDeadline(Promise.all(oldPending), 'Former cumulative observation', 20),
    { code: 'ERR_ACCEPTANCE_OBSERVATION_TIMEOUT' });
});

test('supplemental body observation and cleanup deadlines do not wait forever', async () => {
  await assert.rejects(withDeadline(never(), 'Browser response body', 20),
    { code: 'ERR_ACCEPTANCE_OBSERVATION_TIMEOUT' });
  await assert.rejects(withDeadline(never(), 'Context cleanup', 20),
    { code: 'ERR_ACCEPTANCE_OBSERVATION_TIMEOUT' });
});
