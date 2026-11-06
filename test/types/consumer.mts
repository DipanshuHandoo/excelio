import type { Buffer } from 'node:buffer';
import type { Readable } from 'node:stream';
import excelio, { ExcelIoError, type Spec, type ReadResult, type ParsedWorkbook } from '@dipanshuhandoo/excelio';
import minified from '@dipanshuhandoo/excelio/min';

interface Row { value: number }
const spec: Spec<Row> = [{ workbook: 'typed', sheets: [{ sheet: 'Data', data: [{ value: 1 }] }] }];
const buffers: Promise<Buffer | Buffer[]> = excelio.write(spec);
const target: Promise<string> = excelio.write(spec, { to: 'typed.xlsx', useWorker: true });
const stream: Readable = minified.writeStream(spec);
const result: Promise<ReadResult<Row>> = minified.read<Row>(stream, {
  validateRow: row => row.value < 0 ? { message: 'negative', column: 'value' } : null,
  afterRead: row => ({ value: row.value + 1 })
});
const error = new ExcelIoError('test', { code: 'INVALID_SPEC', details: { value: 1 } });
async function* streamedRows(): AsyncGenerator<Row> { yield { value: 1 }; }
const iterableRows: Iterable<Row> = new Set([{ value: 2 }]);
const iterableSpec: Spec<Row> = [{ workbook: 'iterable', sheets: [
  { sheet: 'Async', data: streamedRows(), totalRows: 1 },
  { sheet: 'Sync', data: iterableRows },
  { sheet: 'Empty' }
] }];
const streamed: Promise<string> = excelio.write(iterableSpec, {
  to: 'large.xlsx',
  maxRowsPerSheet: 500_000,
  sharedStrings: 'auto',
  beforeWrite: (row, ctx) => { const part: number | undefined = ctx.sheetPart; void part; return row; }
});
const sourceError = new ExcelIoError('source', { code: 'SOURCE_ERROR' });
result.then(parsed => {
  const workbook: ParsedWorkbook<Row> = parsed.data[0];
  const name: undefined = workbook.workbook;
  const value: number = workbook.sheets[0].data[0].value;
  return { name, value };
});
void buffers; void target; void error; void streamed; void sourceError;