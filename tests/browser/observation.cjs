const OBSERVATION_TIMEOUT_MS = 10000;

function withDeadline(promise, label, timeoutMs = OBSERVATION_TIMEOUT_MS) {
  let timer;
  const deadline = new Promise((_, reject) => {
    timer = setTimeout(() => {
      const error = new Error(`${label} did not settle within ${timeoutMs}ms`);
      error.code = 'ERR_ACCEPTANCE_OBSERVATION_TIMEOUT';
      reject(error);
    }, timeoutMs);
  });
  return Promise.race([promise, deadline]).finally(() => clearTimeout(timer));
}

function hasAcceptedHandshake(observation) {
  return observation.messages.some(message => message.action === 'test' && message.data === true &&
    message.sourceName === 'worker' && message.targetName === 'main');
}

async function waitForConversionOutcome(readState, timeoutMs = 45000, intervalMs = 50) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const state = await withDeadline(readState(), 'Conversion outcome observation',
      Math.min(3000, Math.max(1, deadline - Date.now())));
    if (state.downloadCount > 0) return { kind: 'download', state };
    if (state.success) return { kind: 'success', state };
    if (state.error) return { kind: 'error', state };
    await new Promise(resolve => setTimeout(resolve, Math.min(intervalMs, Math.max(1, deadline - Date.now()))));
  }
  throw new Error(`Conversion produced no error, success or download within ${timeoutMs}ms`);
}

function normalizeWorkerUrl(rawUrl, pageUrl) {
  return { rawUrl, url: new URL(rawUrl, pageUrl).href };
}

function selectAttemptWorker(workers, startIndex, url) {
  const matches = workers.slice(startIndex).filter(worker => worker.url === url);
  if (matches.length !== 1) throw new Error(`Expected one worker from this attempt, observed ${matches.length}`);
  return matches[0];
}

// A worker whose script failed to load may never get an execution context.
// Observe only the specific worker that already passed PDF.js's real handshake.
// No outstanding evaluations from previous attempts are collected or awaited.
async function observeAcceptedWorker(worker, observation, timeoutMs = OBSERVATION_TIMEOUT_MS) {
  if (!hasAcceptedHandshake(observation)) {
    return { status: 'not-requested', reason: 'No accepted native PDF.js handshake' };
  }
  try {
    const scope = await withDeadline(worker.evaluate(() => ({
      dedicated: typeof DedicatedWorkerGlobalScope !== 'undefined' && self instanceof DedicatedWorkerGlobalScope,
      documentType: typeof document, origin: location.origin,
    })), 'Accepted worker scope observation', timeoutMs);
    return { status: 'fulfilled', scope };
  } catch (error) {
    return { status: error.code === 'ERR_ACCEPTANCE_OBSERVATION_TIMEOUT' ? 'timed-out' : 'rejected',
      error: String(error) };
  }
}

module.exports = { waitForConversionOutcome, normalizeWorkerUrl, withDeadline, hasAcceptedHandshake, selectAttemptWorker, observeAcceptedWorker };
