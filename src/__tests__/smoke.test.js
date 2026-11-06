/**
 * Smoke test — verifies basic round-trip + key behaviours.
 * Run with: npm test
 */
import assert from 'node:assert/strict';
import { Readable, PassThrough } from 'node:stream';
import { EventEmitter } from 'node:events';
import { createRequire, syncBuiltinESMExports } from 'node:module';
import ExcelJS from 'exceljs';
import { excelio, ExcelIoError } from '../index.js';

let passed = 0;
let failed = 0;

async function test(name, fn) {
  try {
    await fn();
    console.log(`  ✓ ${name}`);
    passed++;
  } catch (err) {
    console.error(`  ✗ ${name}`);
    console.error('   ', err.message);
    failed++;
  }
}

console.log('excelio smoke tests\n');

await test('basic round-trip preserves string and number values', async () => {
  const spec = [{
    workbook: 'test',
    sheets: [{
      sheet: 'Data',
      data: [
        { name: 'Alice', amount: 100 },
        { name: 'Bob',   amount: 250.5 }
      ]
    }]
  }];
  const buf = await excelio.write(spec);
  assert.ok(Buffer.isBuffer(buf), 'expected Buffer');

  const result = await excelio.read(buf);
  assert.equal(result.ok, true);
  assert.equal(result.data[0].sheets[0].data.length, 2);
  assert.equal(result.data[0].sheets[0].data[0].name, 'Alice');
  assert.equal(result.data[0].sheets[0].data[0].amount, 100);
  assert.equal(result.data[0].sheets[0].data[1].amount, 250.5);
});

await test('column metadata sets headers and ordering', async () => {
  const spec = [{
    workbook: 'test',
    sheets: [{
      sheet: 'Data',
      columns: [
        { key: 'amount', header: 'Total Sales' },
        { key: 'name',   header: 'Customer Name' }
      ],
      data: [{ name: 'Alice', amount: 100 }]
    }]
  }];
  const buf = await excelio.write(spec);
  const result = await excelio.read(buf);
  const headers = result.data[0].sheets[0].columns.map(c => c.header);
  assert.deepEqual(headers, ['Total Sales', 'Customer Name']);
});

await test('empty cells read as null by default', async () => {
  const spec = [{
    workbook: 'test',
    sheets: [{
      sheet: 'Data',
      data: [{ name: 'Alice', notes: '' }, { name: 'Bob', notes: 'hi' }]
    }]
  }];
  const buf = await excelio.write(spec);
  const result = await excelio.read(buf);
  assert.equal(result.data[0].sheets[0].data[0].notes, null);
  assert.equal(result.data[0].sheets[0].data[1].notes, 'hi');
});

await test('dates default to ISO strings', async () => {
  const d = new Date('2026-05-02T10:30:00.000Z');
  const spec = [{
    workbook: 'test',
    sheets: [{
      sheet: 'Data',
      columns: [{ key: 'when', type: 'date' }],
      data: [{ when: d }]
    }]
  }];
  const buf = await excelio.write(spec);
  const result = await excelio.read(buf);
  const got = result.data[0].sheets[0].data[0].when;
  assert.equal(typeof got, 'string', 'expected ISO string');
  assert.equal(new Date(got).getTime(), d.getTime());
});

await test('validateRow errors are collected', async () => {
  const spec = [{
    workbook: 'test',
    sheets: [{
      sheet: 'Data',
      data: [{ name: 'Alice', age: 30 }, { name: 'Bob', age: -5 }, { name: 'Eve', age: 25 }]
    }]
  }];
  const buf = await excelio.write(spec);
  const result = await excelio.read(buf, {
    validateRow: (row) => row.age < 0 ? { message: 'age must be non-negative', column: 'age' } : null
  });
  assert.equal(result.ok, false);
  assert.equal(result.errors.length, 1);
  assert.equal(result.errors[0].column, 'age');
  assert.equal(result.data[0].sheets[0].data.length, 2); // bad row not in data
});

await test('multi-line cells preserve newlines', async () => {
  const spec = [{
    workbook: 'test',
    sheets: [{ sheet: 'Data', data: [{ note: 'line one\nline two\nline three' }] }]
  }];
  const buf = await excelio.write(spec);
  const result = await excelio.read(buf);
  assert.equal(result.data[0].sheets[0].data[0].note, 'line one\nline two\nline three');
});

await test('multi-language data round-trips', async () => {
  const spec = [{
    workbook: 'test',
    sheets: [{ sheet: 'Data', data: [
      { lang: 'hindi',   text: 'नमस्ते' },
      { lang: 'chinese', text: '你好' },
      { lang: 'arabic',  text: 'مرحبا' }
    ]}]
  }];
  const buf = await excelio.write(spec);
  const result = await excelio.read(buf);
  const rows = result.data[0].sheets[0].data;
  assert.equal(rows[0].text, 'नमस्ते');
  assert.equal(rows[1].text, '你好');
  assert.equal(rows[2].text, 'مرحبا');
});

await test('multiple sheets in one workbook', async () => {
  const spec = [{
    workbook: 'test',
    sheets: [
      { sheet: 'A', data: [{ x: 1 }] },
      { sheet: 'B', data: [{ y: 2 }] }
    ]
  }];
  const buf = await excelio.write(spec);
  const result = await excelio.read(buf);
  assert.equal(result.data[0].sheets.length, 2);
  assert.equal(result.data[0].sheets[0].sheet, 'A');
  assert.equal(result.data[0].sheets[1].sheet, 'B');
});

await test('multi-workbook spec returns array of buffers', async () => {
  const spec = [
    { workbook: 'one', sheets: [{ sheet: 'X', data: [{ a: 1 }] }] },
    { workbook: 'two', sheets: [{ sheet: 'Y', data: [{ b: 2 }] }] }
  ];
  const result = await excelio.write(spec);
  assert.ok(Array.isArray(result));
  assert.equal(result.length, 2);
  assert.ok(Buffer.isBuffer(result[0]));
});

await test('invalid spec throws ExcelIoError', async () => {
  await assert.rejects(() => excelio.write([]), ExcelIoError);
  await assert.rejects(() => excelio.write([{ workbook: 'x' }]), ExcelIoError); // no sheets
});

await test('sheet name sanitisation truncates long names', async () => {
  const longName = 'a'.repeat(40);
  const spec = [{
    workbook: 'test',
    sheets: [{ sheet: longName, data: [{ x: 1 }] }]
  }];
  const buf = await excelio.write(spec);
  const result = await excelio.read(buf);
  assert.equal(result.data[0].sheets[0].sheet.length, 31);
});

const workerSpec = [{ workbook: 'worker', sheets: [{ sheet: 'Data', data: [{ value: 42 }] }] }];

await test('forced worker write returns a Buffer and can be read in a worker', async () => {
  const buffer = await excelio.write(workerSpec, { useWorker: true });
  assert.ok(Buffer.isBuffer(buffer));
  const result = await excelio.read(buffer, { useWorker: true });
  assert.equal(result.data[0].sheets[0].data[0].value, 42);
});

await test('worker multi-workbook output contains Buffers', async () => {
  const buffers = await excelio.write([...workerSpec, ...workerSpec], { useWorker: true });
  assert.equal(buffers.length, 2);
  assert.ok(buffers.every(buffer => Buffer.isBuffer(buffer)));
});

await test('auto worker thresholds work at the exact boundary', async () => {
  const buffer = await excelio.write(workerSpec, { workerThreshold: { rows: 1 } });
  assert.ok(Buffer.isBuffer(buffer));
  const result = await excelio.read(buffer, { workerThreshold: { bytes: buffer.length } });
  assert.equal(result.stats.totalRows, 1);
});

await test('forced worker stream reads fall back inline', async () => {
  const buffer = await excelio.write(workerSpec, { useWorker: false });
  const result = await excelio.read(Readable.from([buffer]), { useWorker: true });
  assert.equal(result.stats.totalRows, 1);
});

await test('forced worker writes with hooks fall back inline', async () => {
  const buffer = await excelio.write(workerSpec, {
    useWorker: true,
    beforeWrite: row => ({ value: row.value + 1 })
  });
  const result = await excelio.read(buffer, { useWorker: true, validateRow: () => null });
  assert.equal(result.data[0].sheets[0].data[0].value, 43);
});

await test('writable targets and writeStream stay inline', async () => {
  const target = new PassThrough();
  const chunks = [];
  target.on('data', chunk => chunks.push(chunk));
  assert.equal(await excelio.write(workerSpec, { to: target, useWorker: true }), target);
  const result = await excelio.read(Buffer.concat(chunks));
  assert.equal(result.stats.totalRows, 1);
  const streamed = [];
  for await (const chunk of excelio.writeStream(workerSpec, { useWorker: true })) streamed.push(chunk);
  assert.equal((await excelio.read(Buffer.concat(streamed))).stats.totalRows, 1);
});

await test('worker parse failures preserve structured errors', async () => {
  await assert.rejects(() => excelio.read(Buffer.from('not xlsx'), { useWorker: true }),
    err => err instanceof ExcelIoError && err.code === 'IO_ERROR');
  await assert.rejects(() => excelio.write([], { useWorker: true }),
    err => err instanceof ExcelIoError && err.code === 'INVALID_SPEC');
  await assert.rejects(() => excelio.write([null]),
    err => err instanceof ExcelIoError && err.code === 'INVALID_SPEC');
});

await test('read options preserve untrimmed text and empty strings', async () => {
  const spec = [{ workbook: 'options', sheets: [{ sheet: 'Data', data: [{ text: ' padded ', blank: '' }] }] }];
  const buffer = await excelio.write(spec);
  const result = await excelio.read(buffer, { nullForBlank: false, trimStrings: false });
  assert.deepEqual(result.data[0].sheets[0].data[0], { text: ' padded ', blank: '' });
});

await test('date output modes survive worker serialization', async () => {
  const date = new Date('2026-01-01T00:00:00.000Z');
  const spec = [{ workbook: 'dates', sheets: [{ sheet: 'Data', columns: [{ key: 'date', type: 'date' }], data: [{ date: date.toISOString() }] }] }];
  const buffer = await excelio.write(spec, { useWorker: true });
  const dates = await excelio.read(buffer, { useWorker: true, dateFormat: 'date' });
  assert.ok(dates.data[0].sheets[0].data[0].date instanceof Date);
  assert.equal(dates.data[0].sheets[0].data[0].date.getTime(), date.getTime());
  const epochs = await excelio.read(buffer, { useWorker: true, dateFormat: 'epoch' });
  assert.equal(epochs.data[0].sheets[0].data[0].date, date.getTime());
});

await test('afterRead runs before validation and fail mode reports row details', async () => {
  const buffer = await excelio.write(workerSpec);
  const result = await excelio.read(buffer, {
    afterRead: row => ({ value: row.value + 1 }),
    validateRow: row => row.value === 43 ? null : { message: 'wrong value' }
  });
  assert.equal(result.ok, true);
  assert.equal(result.data[0].sheets[0].data[0].value, 43);
  await assert.rejects(() => excelio.read(buffer, {
    validateRow: () => ({ message: 'bad row', column: 'value' }), onRowError: 'fail'
  }), error => error instanceof ExcelIoError && error.code === 'ROW_VALIDATION' && error.details.row === 2);
  const failedTransform = await excelio.read(buffer, { afterRead: () => { throw new Error('transform failed'); } });
  assert.equal(failedTransform.errors[0].code, 'TRANSFORM_ERROR');
  assert.equal(failedTransform.stats.totalRows, 0);
});

await test('progress callbacks and pre-aborted signals force inline operation', async () => {
  const writes = [];
  const buffer = await excelio.write(workerSpec, { useWorker: true, batchSize: 1, onProgress: event => writes.push(event) });
  assert.equal(writes.at(-1).rowsProcessed, 1);
  assert.equal(writes.at(-1).phase, 'write');
  const reads = [];
  await excelio.read(buffer, { useWorker: true, batchSize: 1, onProgress: event => reads.push(event) });
  assert.equal(reads[0].rowsProcessed, 1);
  assert.equal(reads[0].phase, 'read');
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(() => excelio.read(buffer, { signal: controller.signal }),
    error => error instanceof ExcelIoError && error.code === 'ABORTED');
  await assert.rejects(() => excelio.write(workerSpec, { signal: controller.signal }),
    error => error instanceof ExcelIoError && error.code === 'ABORTED');
});

await test('streamed workbook retains styling, coercion, and sanitized duplicate names', async () => {
  const sheet = {
    sheet: 'a/b',
    columns: [{ key: 'amount', type: 'number', format: '#,##0.00' }, { key: 'active', type: 'boolean' }],
    styling: { headerBold: true, freezeHeader: true, autoFilter: true },
    data: [{ amount: '12.50', active: 'true' }]
  };
  const buffer = await excelio.write([{ workbook: 'styles', properties: { author: 'tester' }, sheets: [sheet, sheet] }]);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  assert.equal(workbook.creator, 'tester');
  assert.deepEqual(workbook.worksheets.map(worksheet => worksheet.name), ['a_b', 'a_b (2)']);
  const worksheet = workbook.worksheets[0];
  assert.equal(worksheet.getRow(1).font.bold, true);
  assert.equal(worksheet.views[0].state, 'frozen');
  assert.ok(worksheet.autoFilter);
  assert.equal(worksheet.getCell('A2').value, 12.5);
  assert.equal(worksheet.getCell('A2').numFmt, '#,##0.00');
  assert.equal(worksheet.getCell('B2').value, true);
});

await test('non-cloneable worker payloads reject with WORKER_ERROR', async () => {
  const spec = [{ workbook: 'invalid', sheets: [{ sheet: 'Data', data: [{ value: () => 1 }] }] }];
  await assert.rejects(() => excelio.write(spec, { useWorker: true }),
    error => error instanceof ExcelIoError && error.code === 'WORKER_ERROR');
});

await test('a worker exiting successfully without a result rejects', async () => {
  const workerThreads = createRequire(import.meta.url)('node:worker_threads');
  const OriginalWorker = workerThreads.Worker;
  class ExitingWorker extends EventEmitter {
    constructor() {
      super();
      queueMicrotask(() => this.emit('exit', 0));
    }
  }
  try {
    workerThreads.Worker = ExitingWorker;
    syncBuiltinESMExports();
    await assert.rejects(() => excelio.write(workerSpec, { useWorker: true }),
      error => error instanceof ExcelIoError && error.code === 'WORKER_ERROR');
  } finally {
    workerThreads.Worker = OriginalWorker;
    syncBuiltinESMExports();
  }
});

/* ────── 1.1.0: iterable row sources, rollover, shared strings, cleanup ────── */

const sleepMs = (ms) => new Promise(resolve => setTimeout(resolve, ms));
const columns3 = [{ key: 'id', type: 'number' }, { key: 'name' }];
const rowsOf = (n, offset = 0) => Array.from({ length: n }, (_, i) => ({ id: i + 1 + offset, name: `row ${i + 1 + offset}` }));
async function* asyncRows(rows, { delay = 0, onFinally } = {}) {
  try {
    for (const row of rows) {
      if (delay) await sleepMs(delay);
      yield row;
    }
  } finally {
    onFinally?.();
  }
}
const sheetsOf = async (buffer) => (await excelio.read(buffer)).data[0].sheets;
async function worksheetsOf(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook.worksheets;
}
const isCode = (code) => (error) => error instanceof ExcelIoError && error.code === code;

/** Runs fn, then asserts no unhandled rejection or uncaught exception surfaced afterwards. */
async function withoutStrayErrors(fn) {
  const recorded = [];
  const record = (error) => recorded.push(error);
  process.on('unhandledRejection', record);
  process.on('uncaughtException', record);
  try {
    await fn();
    await sleepMs(50);
    assert.deepEqual(recorded.map(error => String(error?.message ?? error)), []);
  } finally {
    process.off('unhandledRejection', record);
    process.off('uncaughtException', record);
  }
}

await test('async generator with explicit columns round-trips', async () => {
  const buffer = await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', columns: columns3, data: asyncRows(rowsOf(3)) }] }]);
  const [sheet] = await sheetsOf(buffer);
  assert.deepEqual(sheet.columns.map(c => c.header), ['id', 'name']);
  assert.deepEqual(sheet.data, rowsOf(3));
});

await test('sync generator round-trips', async () => {
  function* rows() { yield* rowsOf(3); }
  const buffer = await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', columns: columns3, data: rows() }] }]);
  assert.deepEqual((await sheetsOf(buffer))[0].data, rowsOf(3));
});

await test('iterable without columns infers keys from the first row and keeps it', async () => {
  const buffer = await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: asyncRows(rowsOf(3)) }] }]);
  const [sheet] = await sheetsOf(buffer);
  assert.deepEqual(sheet.columns.map(c => c.key), ['id', 'name']);
  assert.equal(sheet.data.length, 3);
  assert.deepEqual(sheet.data[0], { id: 1, name: 'row 1' });
});

await test('object-mode Readable is accepted as row source', async () => {
  const buffer = await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: Readable.from(rowsOf(4), { objectMode: true }) }] }]);
  assert.deepEqual((await sheetsOf(buffer))[0].data, rowsOf(4));
});

await test('empty iterable with columns writes the header only', async () => {
  const buffer = await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', columns: columns3, data: asyncRows([]) }] }]);
  const worksheet = (await worksheetsOf(buffer))[0];
  assert.equal(worksheet.getCell('A1').value, 'id');
  assert.equal((await sheetsOf(buffer))[0].data.length, 0);
});

await test('empty iterable without columns writes an empty sheet', async () => {
  const buffer = await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: asyncRows([]) }] }]);
  const worksheets = await worksheetsOf(buffer);
  assert.equal(worksheets.length, 1);
  assert.equal(worksheets[0].name, 'Data');
  assert.equal(worksheets[0].actualRowCount, 0);
});

await test('array, iterable and array sheets keep order and content', async () => {
  const buffer = await excelio.write([{ workbook: 'it', sheets: [
    { sheet: 'A', data: rowsOf(2) },
    { sheet: 'B', data: asyncRows(rowsOf(3, 10)) },
    { sheet: 'C', data: rowsOf(1, 20) }
  ] }]);
  const sheets = await sheetsOf(buffer);
  assert.deepEqual(sheets.map(s => s.sheet), ['A', 'B', 'C']);
  assert.deepEqual(sheets.map(s => s.data), [rowsOf(2), rowsOf(3, 10), rowsOf(1, 20)]);
});

await test('maxRowsPerSheet rolls an iterable over to continuation sheets', async () => {
  const buffer = await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', columns: columns3, data: asyncRows(rowsOf(7)) }] }],
    { maxRowsPerSheet: 3 });
  const sheets = await sheetsOf(buffer);
  assert.deepEqual(sheets.map(s => s.sheet), ['Data', 'Data (2)', 'Data (3)']);
  assert.deepEqual(sheets.map(s => s.data.length), [3, 3, 1]);
  assert.ok(sheets.every(s => s.columns.map(c => c.header).join() === 'id,name'));
  assert.deepEqual(sheets.flatMap(s => s.data), rowsOf(7));
});

await test('maxRowsPerSheet rolls an array over to continuation sheets', async () => {
  const buffer = await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', columns: columns3, data: rowsOf(7) }] }],
    { maxRowsPerSheet: 3 });
  const sheets = await sheetsOf(buffer);
  assert.deepEqual(sheets.map(s => s.sheet), ['Data', 'Data (2)', 'Data (3)']);
  assert.deepEqual(sheets.map(s => s.data.length), [3, 3, 1]);
});

await test('continuation sheets keep styling', async () => {
  const buffer = await excelio.write([{ workbook: 'it', sheets: [{
    sheet: 'Data', columns: columns3, styling: { headerBold: true, freezeHeader: true, autoFilter: true }, data: asyncRows(rowsOf(4))
  }] }], { maxRowsPerSheet: 3 });
  const worksheets = await worksheetsOf(buffer);
  assert.equal(worksheets.length, 2);
  for (const worksheet of worksheets) {
    assert.equal(worksheet.views[0].state, 'frozen');
    assert.ok(worksheet.autoFilter);
    assert.equal(worksheet.getRow(1).font.bold, true);
    assert.equal(worksheet.getCell('B1').value, 'name');
  }
});

await test('continuation names never take a later logical sheet name', async () => {
  const buffer = await excelio.write([{ workbook: 'it', sheets: [
    { sheet: 'Data', data: rowsOf(5) },
    { sheet: 'Data (2)', data: rowsOf(1, 100) }
  ] }], { maxRowsPerSheet: 3 });
  const sheets = await sheetsOf(buffer);
  assert.deepEqual(sheets.map(s => s.sheet), ['Data', 'Data (3)', 'Data (2)']);
  assert.deepEqual(sheets.map(s => s.data.length), [3, 2, 1]);
});

await test('array specs keep 1.0.0 sheet names', async () => {
  const buffer = await excelio.write([{ workbook: 'it', sheets: [
    { sheet: 'A', data: rowsOf(1) }, { sheet: 'A', data: rowsOf(1) }, { sheet: 'B', data: rowsOf(1) }
  ] }]);
  assert.deepEqual((await worksheetsOf(buffer)).map(w => w.name), ['A', 'A (2)', 'B']);
});

await test('beforeWrite context reports logical index, part and per-part row number', async () => {
  const contexts = [];
  await excelio.write([{ workbook: 'it', sheets: [
    { sheet: 'First', data: rowsOf(1) },
    { sheet: 'Data', columns: columns3, data: asyncRows(rowsOf(5)) }
  ] }], {
    maxRowsPerSheet: 3,
    beforeWrite: (row, ctx) => { contexts.push(ctx); return row; }
  });
  const data = contexts.slice(1);
  assert.equal(data.length, 5);
  assert.deepEqual(data.map(c => c.sheetIndex), [1, 1, 1, 1, 1]);
  assert.deepEqual(data.map(c => c.sheetPart), [1, 1, 1, 2, 2]);
  assert.deepEqual(data.map(c => c.rowNumber), [2, 3, 4, 2, 3]);
  assert.deepEqual(data.map(c => c.sheetName), ['Data', 'Data', 'Data', 'Data (2)', 'Data (2)']);
  assert.equal(contexts[0].sheetPart, 1);
  assert.equal(contexts[0].rowNumber, 2);
});

await test('progress events for iterables carry sheet details and the totalRows hint', async () => {
  const events = [];
  await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', columns: columns3, totalRows: 5, data: asyncRows(rowsOf(5)) }] }],
    { batchSize: 2, maxRowsPerSheet: 3, onProgress: event => events.push(event) });
  assert.deepEqual(events.map(e => e.rowsProcessed), [2, 4, 5]);
  assert.ok(events.every(e => e.totalRows === 5 && e.sheetIndex === 0 && typeof e.sheetName === 'string'));
  assert.deepEqual(events.map(e => [e.sheetName, e.sheetPart]), [['Data', 1], ['Data (2)', 2], ['Data (2)', 2]]);
});

await test('progress totalRows is undefined for an iterable without a hint', async () => {
  const events = [];
  await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: asyncRows(rowsOf(3)) }] }],
    { batchSize: 1, onProgress: event => events.push(event) });
  assert.ok(events.length >= 3);
  assert.ok(events.every(e => e.totalRows === undefined));
  assert.equal(events.at(-1).rowsProcessed, 3);
});

await test('a throwing source rejects with SOURCE_ERROR and row details', async () => {
  const original = new Error('cursor died');
  async function* failing() { yield* rowsOf(2); throw original; }
  await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: failing() }] }]), (error) => {
    assert.equal(error.code, 'SOURCE_ERROR');
    assert.equal(error.cause, original);
    assert.deepEqual(error.details, { workbookIndex: 0, sheetIndex: 0, sheetName: 'Data', rowsRead: 2 });
    return true;
  });
});

await test('a failing iterator factory rejects with SOURCE_ERROR', async () => {
  const source = { [Symbol.asyncIterator]() { throw new Error('cannot open'); } };
  await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: source }] }]), isCode('SOURCE_ERROR'));
});

await test('a non-object iterator result rejects with SOURCE_ERROR', async () => {
  const source = { [Symbol.asyncIterator]() { return { next: async () => 42 }; } };
  await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: source }] }]), isCode('SOURCE_ERROR'));
});

await test('abort during an iterable sheet rejects ABORTED and closes the source', async () => {
  let closed = false;
  const controller = new AbortController();
  setTimeout(() => controller.abort(), 20);
  await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{
    sheet: 'Data', data: asyncRows(rowsOf(10_000), { delay: 1, onFinally: () => { closed = true; } })
  }] }], { signal: controller.signal, batchSize: 1 }), isCode('ABORTED'));
  assert.equal(closed, true);
});

await test('a source failing because of the abort reports ABORTED, not SOURCE_ERROR', async () => {
  const controller = new AbortController();
  const { setTimeout: wait } = await import('node:timers/promises');
  async function* source() { yield rowsOf(1)[0]; await wait(10_000, undefined, { signal: controller.signal }); }
  setTimeout(() => controller.abort(), 20);
  await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: source() }] }], { signal: controller.signal }),
    (error) => error.code === 'ABORTED' && error.cause?.name === 'AbortError');
});

await test('a throwing beforeWrite rejects with that error and closes the source', async () => {
  let closed = false;
  const failure = new Error('bad row');
  await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{
    sheet: 'Data', data: asyncRows(rowsOf(5), { onFinally: () => { closed = true; } })
  }] }], { beforeWrite: (row) => { if (row.id === 3) throw failure; return row; } }), (error) => error === failure);
  assert.equal(closed, true);
});

const { mkdtemp, rm, rename, access } = await import('node:fs/promises');
const { tmpdir } = await import('node:os');
const path = (await import('node:path')).default;
async function* failAfter(n) { yield* rowsOf(n); throw new Error('source failed'); }

await test('failed file write removes the partial file and releases the handle', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'excelio-fail-'));
  const renamed = `${directory}-renamed`;
  try {
    const file = path.join(directory, 'out.xlsx');
    await withoutStrayErrors(async () => {
      await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: failAfter(3) }] }], { to: file }),
        isCode('SOURCE_ERROR'));
    });
    await assert.rejects(() => access(file));
    await rename(directory, renamed);
  } finally {
    await rm(directory, { recursive: true, force: true });
    await rm(renamed, { recursive: true, force: true });
  }
});

await test('failed Buffer write destroys the internal collector', async () => {
  const destroyed = [];
  const originalDestroy = PassThrough.prototype.destroy;
  PassThrough.prototype.destroy = function (...args) { destroyed.push(this); return originalDestroy.apply(this, args); };
  try {
    await withoutStrayErrors(async () => {
      await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: failAfter(3) }] }]), isCode('SOURCE_ERROR'));
    });
  } finally {
    PassThrough.prototype.destroy = originalDestroy;
  }
  assert.ok(destroyed.length >= 1);
});

await test('failed write to a caller stream neither ends nor destroys it', async () => {
  const target = new PassThrough();
  target.resume();
  let finished = false;
  target.on('finish', () => { finished = true; });
  await withoutStrayErrors(async () => {
    await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: failAfter(3) }] }], { to: target }),
      isCode('SOURCE_ERROR'));
  });
  assert.equal(target.destroyed, false);
  assert.equal(finished, false);
  assert.equal(target.listenerCount('error'), 0);
  target.destroy();
});

await test('failed writeStream emits SOURCE_ERROR', async () => {
  await withoutStrayErrors(async () => {
    const stream = excelio.writeStream([{ workbook: 'it', sheets: [{ sheet: 'Data', data: failAfter(3) }] }]);
    stream.resume();
    const error = await new Promise(resolve => stream.once('error', resolve));
    assert.equal(error.code, 'SOURCE_ERROR');
  });
});

await test('an early sink failure interrupts an endless source', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'excelio-sink-'));
  let closed = false;
  async function* endless() {
    try {
      for (let i = 0; ; i++) { await sleepMs(1); yield { id: i }; }
    } finally { closed = true; }
  }
  try {
    const started = Date.now();
    await withoutStrayErrors(async () => {
      await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: endless() }] }],
        { to: path.join(directory, 'missing', 'x.xlsx'), batchSize: 10 }), (error) => error.code === 'ENOENT');
    });
    assert.ok(Date.now() - started < 2000, 'expected a prompt rejection');
    assert.equal(closed, true);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

await test('useWorker: true with iterable data runs inline and warns once', async () => {
  const warnings = [];
  const buffer = await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: asyncRows(rowsOf(3)) }] }],
    { useWorker: true, logger: { warn: (message) => warnings.push(message) } });
  assert.equal((await sheetsOf(buffer))[0].data.length, 3);
  assert.equal(warnings.length, 1);
  const noLogger = await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: asyncRows(rowsOf(2)) }] }], { useWorker: true });
  assert.equal((await sheetsOf(noLogger))[0].data.length, 2);
});

await test('large iterable with useWorker auto runs inline', async () => {
  function* many() { for (let i = 0; i < 60_000; i++) yield { id: i }; }
  const buffer = await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: many() }] }]);
  assert.equal((await excelio.read(buffer)).stats.totalRows, 60_000);
});

await test('invalid data values reject with INVALID_SPEC naming the sheet', async () => {
  for (const data of ['abc', {}, 42, true]) {
    await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Bad', data }] }]),
      (error) => error.code === 'INVALID_SPEC' && error.message.includes('"Bad"'));
  }
});

await test('one iterable used by two sheets is rejected', async () => {
  const shared = asyncRows(rowsOf(2));
  await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{ sheet: 'One', data: shared }, { sheet: 'Two', data: shared }] }]),
    (error) => error.code === 'INVALID_SPEC' && error.message.includes('"Two"') && error.message.includes('"One"'));
});

await test('one array used by two sheets is allowed', async () => {
  const shared = rowsOf(2);
  const buffer = await excelio.write([{ workbook: 'it', sheets: [{ sheet: 'One', data: shared }, { sheet: 'Two', data: shared }] }]);
  assert.deepEqual((await sheetsOf(buffer)).map(s => s.data.length), [2, 2]);
});

await test('invalid maxRowsPerSheet, sharedStrings and batchSize reject with INVALID_SPEC', async () => {
  const spec = [{ workbook: 'it', sheets: [{ sheet: 'Data', data: rowsOf(1) }] }];
  for (const maxRowsPerSheet of [0, 1.5, 2_000_000]) {
    await assert.rejects(() => excelio.write(spec, { maxRowsPerSheet }), isCode('INVALID_SPEC'));
  }
  await assert.rejects(() => excelio.write(spec, { sharedStrings: 'yes' }), isCode('INVALID_SPEC'));
  for (const batchSize of [0, -1, 1.5, Infinity]) {
    await assert.rejects(() => excelio.write(spec, { batchSize }), isCode('INVALID_SPEC'));
  }
  assert.throws(() => excelio.writeStream(spec, { batchSize: 0 }), isCode('INVALID_SPEC'));
});

await test('invalid specs fail before worker routing', async () => {
  const workerThreads = createRequire(import.meta.url)('node:worker_threads');
  const OriginalWorker = workerThreads.Worker;
  let spawned = 0;
  workerThreads.Worker = class extends EventEmitter { constructor() { super(); spawned++; } };
  syncBuiltinESMExports();
  try {
    await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: 'abc' }] }], { useWorker: true }),
      isCode('INVALID_SPEC'));
    await assert.rejects(() => excelio.write([{ workbook: 'it', sheets: [{ sheet: 'Data', data: rowsOf(1) }] }], { useWorker: true, maxRowsPerSheet: 0 }),
      isCode('INVALID_SPEC'));
    assert.equal(spawned, 0);
  } finally {
    workerThreads.Worker = OriginalWorker;
    syncBuiltinESMExports();
  }
});

await test('sharedStrings modes control the cell string type', async () => {
  const JSZip = createRequire(import.meta.url)('jszip');
  const inspect = async (buffer) => {
    const zip = await JSZip.loadAsync(buffer);
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');
    return { shared: Boolean(zip.file('xl/sharedStrings.xml')), s: sheetXml.includes('t="s"'), str: sheetXml.includes('t="str"') };
  };
  const arraySpec = () => [{ workbook: 'it', sheets: [{ sheet: 'Data', columns: columns3, data: rowsOf(2) }] }];
  const iterableSpec = () => [{ workbook: 'it', sheets: [{ sheet: 'Data', columns: columns3, data: asyncRows(rowsOf(2)) }] }];
  const cases = [
    [arraySpec, undefined, true],
    [iterableSpec, undefined, false],
    [arraySpec, false, false],
    [iterableSpec, true, true]
  ];
  for (const [makeSpec, sharedStrings, expectShared] of cases) {
    const buffer = await excelio.write(makeSpec(), sharedStrings === undefined ? {} : { sharedStrings });
    const result = await inspect(buffer);
    assert.equal(result.shared, expectShared);
    assert.equal(result.s, expectShared);
    assert.equal(result.str, !expectShared);
    assert.deepEqual((await sheetsOf(buffer))[0].data, rowsOf(2));
  }
});

await test('writeStream accepts an iterable sheet', async () => {
  const chunks = [];
  for await (const chunk of excelio.writeStream([{ workbook: 'it', sheets: [{ sheet: 'Data', data: asyncRows(rowsOf(3)) }] }])) {
    chunks.push(chunk);
  }
  assert.deepEqual((await sheetsOf(Buffer.concat(chunks)))[0].data, rowsOf(3));
});

await test('multi-workbook Buffer output with an iterable in each workbook', async () => {
  const buffers = await excelio.write([
    { workbook: 'one', sheets: [{ sheet: 'A', data: asyncRows(rowsOf(2)) }] },
    { workbook: 'two', sheets: [{ sheet: 'B', data: asyncRows(rowsOf(3)) }] }
  ]);
  assert.equal(buffers.length, 2);
  assert.equal((await sheetsOf(buffers[0]))[0].data.length, 2);
  assert.equal((await sheetsOf(buffers[1]))[0].data.length, 3);
});

console.log(`\n${passed} passed, ${failed} failed\n`);
// No process.exit(): a leftover handle (incomplete cleanup) shows up as a hang.
process.exitCode = failed === 0 ? 0 : 1;
