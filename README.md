# excelio

Plug-and-play Excel ↔ JSON utility. Streams in/out so 1M+ row datasets stay memory-safe. Worker-thread offload above a threshold. BYO validation, BYO logger.

## Install

The directory is self-contained and only depends on `exceljs`. Drop the `excelio/` folder anywhere and `npm install exceljs`.

## API

```js
import { excelio } from './excelio';
```

### Spec shape (input to write, output of read)

```js
[
  {
    workbook: 'sales-2026',
    properties: { author: 'system' },     // optional
    sheets: [
      {
        sheet: 'Q1',
        columns: [                          // optional; inferred from first row if omitted
          { key: 'name',   header: 'Customer',     width: 30 },
          { key: 'amount', header: 'Total',        type: 'number', format: '#,##0.00' },
          { key: 'date',   header: 'Order Date',   type: 'date' }
        ],
        styling: {                          // optional
          headerBold: true,
          freezeHeader: true,
          autoFilter: true
        },
        data: [
          { name: 'Alice', amount: 100, date: new Date('2026-01-15') }
        ]
      }
    ]
  }
]
```

### Write

```js
// Returns Buffer (or array of Buffers when spec has multiple workbooks)
const buf = await excelio.write(spec);

// Write to file
await excelio.write(spec, { to: 'output.xlsx' });

// Pipe to a Writable (HTTP response, S3 upload, etc.)
await excelio.write(spec, { to: response });

// Get a Readable stream you pipe yourself
const stream = excelio.writeStream(spec);
stream.pipe(response);
```

### Read

```js
// From buffer
const result = await excelio.read(buffer);

// From file path
const result = await excelio.read('input.xlsx');

// From stream
const result = await excelio.read(req);
```

Result:

```js
{
  ok: true,
  data: [{ workbook, sheets: [{ sheet, columns, data: [...] }] }],
  errors: [
    { workbook: 0, sheet: 'Q1', row: 47, column: 'amount',
      code: 'TYPE_MISMATCH', message: 'expected number' }
  ],
  stats: { totalRows: 12345, totalSheets: 1, durationMs: 1820 }
}
```

`ok: false` when any errors present. Caller decides whether to use `data` partially or abort.

### Validation (BYO)

```js
const result = await excelio.read(buf, {
  validateRow(row, ctx) {
    if (!row.email?.includes('@')) {
      return { code: 'INVALID_EMAIL', message: 'invalid email', column: 'email' };
    }
    return null;          // valid
  },
  onRowError: 'collect'   // default; or 'fail' to throw on first
});
```

`ctx`: `{ workbookIndex, workbookName, sheetIndex, sheetName, rowNumber }`.

### Transform hooks

```js
await excelio.read(buf, {
  afterRead(row, ctx) {
    return { ...row, email: row.email?.toLowerCase() };
  }
});

await excelio.write(spec, {
  beforeWrite(row, ctx) {
    return { ...row, exported_at: new Date() };
  }
});
```

### Progress + cancellation

```js
const ac = new AbortController();

await excelio.read(buf, {
  signal: ac.signal,
  onProgress({ phase, sheetName, rowsProcessed }) {
    console.log(`${phase} ${sheetName}: ${rowsProcessed} rows`);
  }
});

// cancel
ac.abort();   // throws ExcelIoError({ code: 'ABORTED' })
```

### Worker threads

By default, jobs above ~50,000 rows or ~10 MB are offloaded to a worker thread automatically:

```js
await excelio.write(largeSpec);                     // auto-uses worker
await excelio.write(spec, { useWorker: true });     // force worker
await excelio.write(spec, { useWorker: false });    // force inline
await excelio.write(spec, {
  workerThreshold: { rows: 100_000, bytes: 50 * 1024 * 1024 }
});
```

Hooks (`validateRow`, `beforeWrite`, `afterRead`, `onProgress`, `signal`, `logger`) and non-string `to` targets force inline execution because functions and streams can't cross the worker boundary.

### Date handling

- **On write**: accepts `Date`, ISO 8601 strings, and unix ms numbers when `column.type === 'date'`.
- **On read**: by default returns ISO 8601 strings (`'2026-05-02T10:30:00.000Z'`).
  - `dateFormat: 'date'` → returns `Date` objects
  - `dateFormat: 'epoch'` → returns unix ms numbers

### Empty cells

Default: `null`. Override with `nullForBlank: false` to get `''`.

### Header ↔ key round-tripping

The reader keys parsed rows by **header text** (since that's all an `.xlsx` file actually stores about columns). If you write with a `header` that differs from the `key`, reading back will use the header as the property name:

```js
// Write
columns: [{ key: 'firstName', header: 'First Name' }]
// Read back
data: [{ 'First Name': 'Alice' }]   // not { firstName: ... }
```

For round-trippable data, omit `header` (it defaults to `key`) or use the same string for both. Use distinct headers only when producing reports for human consumption that won't be re-parsed.

### Sheet name sanitisation

Excel limits sheet names to 31 chars and forbids `\ / ? * [ ] :`. The utility auto-truncates and replaces forbidden chars; duplicate names get suffixed `(2)`, `(3)`. Pass `logger.warn` to be notified when this happens.

### Multi-language

Full UTF-8/16. Hindi, Chinese, Arabic, etc. round-trip cleanly.

## Limitations

- **Multi-workbook + `to`**: not supported. Either omit `to` to get an array of buffers, or call `excelio.write` per workbook.
- **Charts, pivot tables, conditional formatting, macros, encryption**: out of scope. Use `exceljs` directly if needed.
- **Formula cells on read**: returns the cached result, not the formula text.

## Errors

```js
import { ExcelIoError } from './excelio';

// thrown on:
// - INVALID_SPEC    — bad input
// - IO_ERROR        — corrupt file, write failure
// - ABORTED         — signal.aborted
// - WORKER_ERROR    — worker crashed
// - ROW_VALIDATION  — only when onRowError: 'fail'
```

Per-row validation errors don't throw (default) — they accumulate in `result.errors[]`.
