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

console.log(`\n${passed} passed, ${failed} failed\n`);
process.exit(failed === 0 ? 0 : 1);
