// This protocol is deliberately limited to the included synthetic fixtures.
// It observes real canvas identities, not engine-specific stack function names.
// Keep the factory self-contained: the same function is injected into the page.
function createCanvasBoundary(evidence) {
  const identities = new WeakMap();
  let nextId = 0;
  let preceding = null;
  const compactStack = stack => String(stack || '').split('\n').slice(0, 7).join('\n').slice(0, 1800);
  const diagnose = (boundary, canvas, stack, reason) => {
    if (evidence.unmatched.length < 12) evidence.unmatched.push({ boundary, reason,
      width: canvas.width, height: canvas.height, stack: compactStack(stack) });
  };
  const violation = (boundary, identity, reason) => {
    if (evidence.protocolViolations.length < 20) evidence.protocolViolations.push({
      boundary, canvasId: identity?.id ?? null, attempt: evidence.attempt, reason });
  };
  const fixtureSize = canvas => evidence.fixture?.pages.some(page =>
    canvas.width === page.widthPt * 1.5 && canvas.height === page.heightPt * 1.5);
  return {
    register(canvas) {
      identities.set(canvas, { id: ++nextId, createdAttempt: evidence.attempt,
        contextRequests: 0, exports: 0, assignment: null });
    },
    context(canvas, args, stack) {
      const identity = identities.get(canvas);
      if (identity) identity.contextRequests++;
      if (!evidence.fixture || args[0] !== '2d') return null;
      if (identity?.assignment?.attempt === evidence.attempt) {
        violation('getContext', identity, 'duplicate context request on a source canvas');
        return null;
      }
      if (!identity || !fixtureSize(canvas) || canvas.isConnected || args.length !== 1) {
        diagnose('getContext', canvas, stack, 'outside the fixture source-canvas protocol');
        return null;
      }
      if (identity.createdAttempt !== evidence.attempt || identity.contextRequests !== 1) {
        violation('getContext', identity, 'source candidate was reused or had an earlier context request');
        return null;
      }
      const page = evidence.canvases + 1;
      if (page > evidence.fixture.pageCount) {
        violation('getContext', identity, 'extra source-sized canvas after the final page');
        return null;
      }
      if (preceding?.assignment?.attempt === evidence.attempt && preceding.exports !== 1) {
        violation('getContext', identity, 'source-sized canvas interleaved before preceding page export');
        return null;
      }
      const expected = evidence.fixture.pages[page - 1];
      if (canvas.width !== expected.widthPt * 1.5 || canvas.height !== expected.heightPt * 1.5) {
        violation('getContext', identity, 'source page viewport is out of order');
        return null;
      }
      identity.assignment = { page, attempt: evidence.attempt };
      preceding = identity;
      evidence.canvases++;
      const observed = { page, invocation: evidence.canvases, canvasId: identity.id,
        createdAttempt: identity.createdAttempt, width: canvas.width, height: canvas.height,
        arguments: ['2d'], stack: compactStack(stack) };
      evidence.contextCalls.push(observed);
      return observed;
    },
    export(canvas, args, stack) {
      const identity = identities.get(canvas);
      const assignment = identity?.assignment;
      if (!assignment || assignment.attempt !== evidence.attempt) {
        if (fixtureSize(canvas) && args[0] === 'image/jpeg' && args[1] === 0.95) {
          violation('toDataURL', identity, 'source-like export without a current source-canvas context');
        }
        diagnose('toDataURL', canvas, stack, 'no current source-canvas identity');
        return null;
      }
      if (identity !== preceding || identity.exports !== 0 || assignment.page !== evidence.exportCalls + 1) {
        violation('toDataURL', identity, 'duplicate or out-of-order source-page export');
        return null;
      }
      const expected = evidence.fixture.pages[assignment.page - 1];
      if (args.length !== 2 || args[0] !== 'image/jpeg' || args[1] !== 0.95 ||
          canvas.width !== expected.widthPt * 1.5 || canvas.height !== expected.heightPt * 1.5) {
        violation('toDataURL', identity, 'source export arguments or viewport changed');
        return null;
      }
      identity.exports++;
      return { page: assignment.page, invocation: ++evidence.exportCalls, canvasId: identity.id,
        arguments: ['image/jpeg', 0.95], stack: compactStack(stack) };
    },
  };
}

module.exports = { createCanvasBoundary };
