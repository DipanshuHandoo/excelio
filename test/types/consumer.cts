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
void buffers; void target; void result; void error;