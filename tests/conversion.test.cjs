const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createRequire } = require('node:module');
const { test } = require('node:test');
const ts = require('typescript');
const docx = require('docx');
const JSZip = createRequire(require.resolve('docx'))('jszip');

// Execute the actual component handler, with adapters for React state, PDF.js,
// canvas and downloads. DOCX construction, packing and ZIP/media checks are real.
// Browser rendering and the same-origin PDF.js worker need separate acceptance QA.
const componentPath = path.join(__dirname, '../src/components/PdfToWordConverter.tsx');
const compiled = ts.transpileModule(fs.readFileSync(componentPath, 'utf8'), {
  compilerOptions: {
    module: ts.ModuleKind.CommonJS,
    jsx: ts.JsxEmit.ReactJSX,
    target: ts.ScriptTarget.ES2020,
    esModuleInterop: true,
  },
}).outputText;
const sourcePages = [
  { width: 6, height: 9 },
  { width: 9, height: 6 },
  { width: 6, height: 6 },
].map((size, index) => ({
  ...size,
  bytes: fs.readFileSync(path.join(__dirname, `fixtures/page-${index + 1}.jpg`)),
}));

function findElement(node, type) {
  if (!node || typeof node !== 'object') return undefined;
  if (node.type === type) return node;
  const children = node.props?.children;
  for (const child of Array.isArray(children) ? children.flat(Infinity) : [children]) {
    const found = findElement(child, type);
    if (found) return found;
  }
}

function createHarness(initialFailure, failurePage = 2) {
  const states = [];
  const events = [];
  const downloads = [];
  let hookIndex = 0;
  let currentPage = 0;
  let failure = initialFailure;
  let imageCount = 0;
  let packCount = 0;
  let beforePack;

  const react = {
    useState(initial) {
      const index = hookIndex++;
      if (!(index in states)) states[index] = initial;
      return [states[index], update => {
        states[index] = typeof update === 'function' ? update(states[index]) : update;
        if (index === 2) events.push({ type: 'progress', value: states[index] });
      }];
    },
  };
  const pdfjs = {
    GlobalWorkerOptions: {},
    getDocument(options) {
      events.push({ type: 'load', error: states[3], isEvalSupported: options.isEvalSupported });
      imageCount = 0;
      return { promise: Promise.resolve({
        numPages: sourcePages.length,
        async getPage(pageNumber) {
          currentPage = pageNumber;
          events.push({ type: 'page', page: pageNumber });
          if (failure === 'load-page' && pageNumber === failurePage) throw new Error('page unavailable');
          return {
            getViewport: () => sourcePages[pageNumber - 1],
            render() {
              events.push({ type: 'render', page: pageNumber });
              return { promise: failure === 'render' && pageNumber === failurePage
                ? Promise.reject(new Error('render failed')) : Promise.resolve() };
            },
          };
        },
      }) };
    },
  };
  const context = {
    exports: {}, Uint8Array, ArrayBuffer, Blob, atob, Error, Date,
    console: { error: (...args) => events.push({ type: 'console-error', args }) },
    require(name) {
      if (name === 'react') return react;
      if (name === 'pdfjs-dist') return pdfjs;
      if (name === 'pdfjs-dist/build/pdf.worker.min.js?url') return '/assets/pdf.worker.min.js';
      if (name === 'file-saver') return { saveAs(blob, filename) {
        if (failure === 'save') throw new Error('download failed');
        downloads.push({ blob, filename });
        events.push({ type: 'download' });
      } };
      if (name === 'docx') return {
        ...docx,
        ImageRun: class extends docx.ImageRun {
          constructor(options) {
            imageCount++;
            if (failure === 'insert' && imageCount === failurePage) throw new Error('image insertion failed');
            super(options);
          }
        },
        Packer: { async toBlob(document) {
          packCount++;
          events.push({ type: 'pack' });
          if (beforePack) await beforePack();
          if (failure === 'pack') throw new Error('packing failed');
          return docx.Packer.toBlob(document);
        } },
      };
      return require(name);
    },
    document: { createElement(tag) {
      assert.equal(tag, 'canvas');
      const pageNumber = currentPage;
      return {
        width: 0, height: 0,
        getContext(type) {
          assert.equal(type, '2d');
          return failure === 'context' && pageNumber === failurePage ? null : {};
        },
        toDataURL(type, quality) {
          assert.equal(type, 'image/jpeg');
          assert.equal(quality, 0.95);
          if (pageNumber === failurePage) {
            if (failure === 'encode') throw new Error('encoding failed');
            if (failure === 'empty-image') return 'data:,';
            if (failure === 'malformed-base64') return 'data:image/jpeg;base64,%';
            if (failure === 'empty-bytes') return 'data:image/jpeg;base64, ';
          }
          return `data:image/jpeg;base64,${sourcePages[pageNumber - 1].bytes.toString('base64')}`;
        },
      };
    } },
  };
  vm.runInNewContext(compiled, context, { filename: componentPath });
  function render() {
    hookIndex = 0;
    return context.exports.default();
  }
  const file = { name: 'three-pages.pdf', size: 100, type: 'application/pdf',
    arrayBuffer: async () => new ArrayBuffer(0) };
  findElement(render(), 'input').props.onChange({ target: { files: [file] } });

  return {
    events, downloads,
    get packCount() { return packCount; },
    get state() {
      return { isConverting: states[1], progress: states[2], error: states[3], log: states[4] };
    },
    get button() { return findElement(render(), 'button'); },
    convert: () => findElement(render(), 'button').props.onClick(),
    setFailure: next => { failure = next; },
    beforePack: callback => { beforePack = callback; },
  };
}

function hasSuccess(harness) {
  return harness.state.log.some(line => line.includes('Chuyển đổi hoàn tất!'));
}

async function assertCompleteDocument(download) {
  assert.equal(download.filename, 'three-pages_converted.docx');
  const zip = await JSZip.loadAsync(await download.blob.arrayBuffer());
  const xml = await zip.file('word/document.xml').async('string');
  const rels = await zip.file('word/_rels/document.xml.rels').async('string');
  const imageIds = [...xml.matchAll(/<a:blip[^>]*\br:embed="([^"]+)"/g)].map(match => match[1]);
  const relationships = new Map([...rels.matchAll(/<Relationship\b[^>]*\/>/g)].map(([tag]) => {
    return [/\bId="([^"]+)"/.exec(tag)[1], /\bTarget="([^"]+)"/.exec(tag)[1]];
  }));
  assert.equal(imageIds.length, sourcePages.length, 'one embedded image per PDF page');
  const media = Object.keys(zip.files).filter(name => name.startsWith('word/media/') && !zip.files[name].dir);
  assert.equal(media.length, sourcePages.length, 'no missing or extra media');
  for (const [index, id] of imageIds.entries()) {
    const target = path.posix.normalize(path.posix.join('word', relationships.get(id)));
    assert.deepEqual(await zip.file(target).async('nodebuffer'), sourcePages[index].bytes,
      `image ${index + 1} must contain the corresponding source page bytes, in order`);
  }
  assert.match(xml, /Tài liệu chuyển đổi từ: three-pages\.pdf/);
  assert.match(xml, /Được tạo vào:/);
  assert.deepEqual([...xml.matchAll(/<w:t[^>]*>(Trang \d+)<\/w:t>/g)].map(match => match[1]),
    ['Trang 1', 'Trang 2', 'Trang 3']);
  const paragraphs = xml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g);
  assert.equal(paragraphs.length, 12, 'retain title, date, spacer and three paragraphs per page');
  for (let index = 0; index < sourcePages.length; index++) {
    assert.match(paragraphs[3 + index * 3], new RegExp(`Trang ${index + 1}`));
    assert.match(paragraphs[4 + index * 3], /<a:blip\b/);
    assert.match(paragraphs[5 + index * 3], /w:after="200"/);
  }
  assert.deepEqual([...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"/g)].map(match => match.slice(1)),
    sourcePages.map(page => ['5715000', String(Math.floor(600 * page.height / page.width) * 9525)]),
    'preserve the 600px image width and each source page aspect ratio');
}

test('healthy conversion embeds every source page in order and preserves Word layout', async () => {
  const harness = createHarness();
  harness.beforePack(() => {
    assert.equal(harness.state.progress, 99);
    assert.equal(harness.state.isConverting, true);
    assert.equal(harness.button.props.disabled, true);
    assert.equal(hasSuccess(harness), false);
    assert.equal(harness.downloads.length, 0);
  });
  await harness.convert();
  assert.equal(harness.packCount, 1);
  assert.equal(harness.downloads.length, 1);
  await assertCompleteDocument(harness.downloads[0]);
  assert.deepEqual(harness.events.filter(event => event.type === 'progress').map(event => event.value),
    [0, 33, 66, 99, 100]);
  assert.ok(harness.events.findIndex(event => event.type === 'download') <
    harness.events.findIndex(event => event.type === 'progress' && event.value === 100));
  assert.equal(harness.state.error, '');
  assert.deepEqual(harness.events.filter(event => event.type === 'load').map(event => event.isEvalSupported), [false]);
  assert.equal(harness.state.isConverting, false);
  assert.equal(hasSuccess(harness), true);
});

for (const failure of ['context', 'empty-image', 'malformed-base64', 'empty-bytes', 'render', 'encode', 'insert', 'load-page']) {
  test(`page 2 ${failure} failure aborts before packing or downloading`, async () => {
    const harness = createHarness(failure);
    await harness.convert();
    assert.equal(harness.packCount, 0);
    assert.equal(harness.downloads.length, 0);
    assert.equal(hasSuccess(harness), false);
    assert.match(harness.state.error, /Không thể chuyển đổi trang 2:/);
    assert.match(harness.state.log.at(-1), /Đã xảy ra lỗi: Không thể chuyển đổi trang 2:/);
    assert.deepEqual(harness.events.filter(event => event.type === 'page').map(event => event.page), [1, 2]);
    assert.deepEqual(harness.events.filter(event => event.type === 'progress').map(event => event.value), [0, 33]);
    assert.equal(harness.state.isConverting, false);
    assert.equal(harness.button.props.disabled, false, 'allow a deliberate retry');
  });
}

test('last-page rendering failure never reaches packing, download or 100%', async () => {
  const harness = createHarness('render', 3);
  await harness.convert();
  assert.equal(harness.packCount, 0);
  assert.equal(harness.downloads.length, 0);
  assert.equal(hasSuccess(harness), false);
  assert.match(harness.state.error, /Không thể chuyển đổi trang 3: render failed/);
  assert.deepEqual(harness.events.filter(event => event.type === 'page').map(event => event.page), [1, 2, 3]);
  assert.deepEqual(harness.events.filter(event => event.type === 'progress').map(event => event.value), [0, 33, 66]);
  assert.equal(harness.state.progress, 66);
  assert.equal(harness.state.isConverting, false);
  assert.equal(harness.button.props.disabled, false);
});

for (const failure of ['pack', 'save']) {
  test(`${failure} failure never reports a completed conversion`, async () => {
    const harness = createHarness(failure);
    harness.beforePack(() => {
      assert.equal(harness.state.progress, 99);
      assert.equal(harness.state.isConverting, true);
      assert.equal(hasSuccess(harness), false);
    });
    await harness.convert();
    assert.equal(harness.packCount, 1);
    assert.equal(harness.downloads.length, 0);
    assert.equal(hasSuccess(harness), false);
    assert.equal(harness.state.progress, 99);
    assert.equal(harness.events.some(event => event.type === 'progress' && event.value === 100), false);
    assert.match(harness.state.error, new RegExp(failure === 'pack' ? 'packing failed' : 'download failed'));
    assert.equal(harness.state.isConverting, false);
    assert.equal(harness.button.props.disabled, false);
  });
}

test('retry clears the previous error immediately and downloads a complete document', async () => {
  const harness = createHarness('empty-image');
  await harness.convert();
  assert.match(harness.state.error, /trang 2/);
  harness.setFailure(undefined);
  const retry = harness.convert();
  assert.equal(harness.state.error, '', 'clear stale errors before any awaited work');
  assert.equal(harness.state.progress, 0);
  assert.equal(harness.state.isConverting, true);
  await retry;
  assert.equal(harness.state.error, '');
  assert.equal(harness.state.progress, 100);
  assert.equal(hasSuccess(harness), true);
  assert.equal(harness.downloads.length, 1);
  assert.equal(harness.packCount, 1);
  assert.equal(harness.state.log.some(line => line.includes('Đã xảy ra lỗi:')), false);
  await assertCompleteDocument(harness.downloads[0]);
  assert.deepEqual(harness.events.filter(event => event.type === 'load').map(event => event.isEvalSupported), [false, false]);
});
