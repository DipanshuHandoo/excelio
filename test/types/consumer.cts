import type { Buffer } from 'node:buffer';
import { PassThrough } from 'node:stream';
import library = require('@dipanshuhandoo/excelio');
import minified = require('@dipanshuhandoo/excelio/min');

const spec: library.Spec = [{ workbook: 'typed', sheets: [{ sheet: 'Data', data: [{ value: 1 }] }] }];
const buffers: Promise<Buffer | Buffer[]> = library.excelio.write(spec);
const writable = new PassThrough();
const target: Promise<PassThrough> = library.default.write(spec, { to: writable });
const result: Promise<library.ReadResult> = minified.excelio.read(writable);
const error: library.ExcelIoError = new library.ExcelIoError('test', { code: 'WORKER_ERROR' });
function* generated(): Generator<{ value: number }> { yield { value: 1 }; }
const iterableSpec: library.Spec<{ value: number }> = [{ workbook: 'iterable', sheets: [{ sheet: 'Data', data: generated() }, { sheet: 'Empty' }] }];
const streamed: Promise<Buffer | Buffer[]> = library.excelio.write(iterableSpec, {
  maxRowsPerSheet: 10,
  sharedStrings: false,
  beforeWrite: (row, ctx) => (ctx.sheetPart === 1 ? row : undefined)
});
const sourceError = new library.ExcelIoError('source', { code: 'SOURCE_ERROR' });
void buffers; void target; void result; void error; void streamed; void sourceError;