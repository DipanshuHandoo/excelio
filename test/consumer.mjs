import assert from 'node:assert/strict';
import { Buffer } from 'node:buffer';
import { mkdtemp, rm } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { PassThrough, Readable } from 'node:stream';

const format = process.argv[2];
assert.ok(['esm', 'cjs', 'esm-min', 'cjs-min'].includes(format));
const require = createRequire(import.meta.url);
const entry = `@dipanshuhandoo/excelio${format.endsWith('-min') ? '/min' : ''}`;
const library = format.startsWith('cjs') ? require(entry) : await import(entry);
const { excelio, ExcelIoError } = library;
assert.equal(library.default, excelio);
assert.equal(require('@dipanshuhandoo/excelio/package.json').name, '@dipanshuhandoo/excelio');
const spec = [{ workbook: 'installed', sheets: [{ sheet: 'Data', data: [{ value: 42, name: 'Alice' }] }] }];
const temporary = await mkdtemp(path.join(tmpdir(), 'excelio-consumer-'));

try {
  for (const useWorker of [false, true]) {
    const buffer = await excelio.write(spec, { useWorker });
    assert.ok(Buffer.isBuffer(buffer));
    const result = await excelio.read(buffer, { useWorker });
    assert.equal(result.ok, true);
    assert.equal(result.stats.totalRows, 1);
    assert.deepEqual(result.data[0].sheets[0].data[0], { value: 42, name: 'Alice' });
    assert.equal(result.data[0].workbook, undefined);
    const buffers = await excelio.write([...spec, ...spec], { useWorker });
    assert.equal(buffers.length, 2);
    assert.ok(buffers.every(buffer => Buffer.isBuffer(buffer)));
    const file = path.join(temporary, `${useWorker}.xlsx`);
    assert.equal(await excelio.write(spec, { to: file, useWorker }), file);
    assert.equal((await excelio.read(file, { useWorker })).stats.totalRows, 1);
  }

  const buffer = await excelio.write(spec, { workerThreshold: { rows: 1 } });
  assert.ok(Buffer.isBuffer(buffer));
  assert.equal((await excelio.read(buffer, { workerThreshold: { bytes: buffer.length } })).stats.totalRows, 1);
  assert.equal((await excelio.read(Readable.from([buffer]), { useWorker: true })).stats.totalRows, 1);
  const target = new PassThrough();
  const chunks = [];
  target.on('data', chunk => chunks.push(chunk));
  assert.equal(await excelio.write(spec, { to: target, useWorker: true }), target);
  assert.equal((await excelio.read(Buffer.concat(chunks))).stats.totalRows, 1);
  const streamed = [];
  for await (const chunk of excelio.writeStream(spec)) streamed.push(chunk);
  assert.equal((await excelio.read(Buffer.concat(streamed))).stats.totalRows, 1);
  const transformed = await excelio.write(spec, { useWorker: true, beforeWrite: row => ({ ...row, value: 43 }) });
  const validated = await excelio.read(transformed, { useWorker: true, validateRow: row => row.value === 43 ? null : { message: 'wrong value' } });
  assert.equal(validated.ok, true);
  assert.equal(validated.data[0].sheets[0].data[0].value, 43);
  await assert.rejects(() => excelio.read(Buffer.from('invalid'), { useWorker: true }),
    error => error instanceof ExcelIoError && error.code === 'IO_ERROR');
  await assert.rejects(() => excelio.write([], { useWorker: true }),
    error => error instanceof ExcelIoError && error.code === 'INVALID_SPEC');
  console.log(`Installed ${format}: inline/worker round-trips, multi-workbooks, files, streams, hooks, and errors passed.`);
} finally {
  await rm(temporary, { recursive: true, force: true });
}