/**
 * Smoke test — verifies basic round-trip + key behaviours.
 * Run with: node src/shared/utils/excelio/__tests__/smoke.test.js
 */
import assert from 'node:assert/strict';
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

console.log(`\n${passed} passed, ${failed} failed\n`);
process.exit(failed === 0 ? 0 : 1);
