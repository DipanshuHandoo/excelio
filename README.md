# ExcelIO

`@dipanshuhandoo/excelio` converts JavaScript row objects to XLSX workbooks and
reads XLSX files into structured data. Rows can come from arrays or be streamed
from generators, database cursors, and paged APIs with bounded memory. It is a
Node.js library, not a browser library. Its only runtime dependency is ExcelJS.

## Features

**Writing**

- One call writes one or many workbooks, each with multiple sheets, to a Buffer,
  a file path, a writable stream, or a readable stream (`writeStream`).
- Sheet `data` can be an array, a sync iterable, or an async iterable (async
  generator, MongoDB/Mongoose cursor, object-mode stream, paged API).
- Large exports: streamed rows with a file or stream target keep memory flat
  (about 50 MB growth for 1,000,000 rows), with output backpressure built in.
- Automatic sheet rollover at Excel's row limit, or earlier with
  `maxRowsPerSheet`: `Links`, `Links (2)`, `Links (3)` … keep columns, header,
  and styling.
- Columns: explicit keys, headers, widths, types (`string`, `number`, `boolean`,
  `date`), and Excel number formats, or keys inferred from the first row.
- Write-side coercion of numeric/boolean strings and dates; multiline and
  multilingual text.
- Styling: bold header, frozen header row, auto-filter; workbook properties
  (author, company, title, created, modified).
- Safe sheet names: forbidden characters replaced, 31-character limit,
  duplicate suffixes.
- Shared-strings control (`sharedStrings`) to trade file size for memory.

**Reading**

- Reads a Buffer, file path, or readable stream into
  `{ ok, data, errors, stats }`, covering all sheets with headers as keys.
- Options for header row, blank handling, trimming, and date output (ISO string,
  Date, or epoch).
- Formula cells return cached values; rich text and hyperlinks return displayed
  text.

**Transforms and validation**

- `beforeWrite` and `afterRead` row hooks with row context (workbook, sheet,
  part, Excel row number).
- `validateRow` collects per-row errors (partial data accepted) or fails fast
  with `onRowError: 'fail'`.

**Operations**

- Worker-thread offload for large plain jobs (`useWorker: 'auto' | true | false`,
  configurable thresholds); hooks, streams, and iterables run inline
  automatically.
- Progress callbacks (`onProgress`, `batchSize`) with sheet, part, and total row
  information.
- Cancellation through `AbortSignal`; streamed sheets respond during the write.
- Failure cleanup: a failed file write removes the partial file, and a caller's
  stream is not ended or destroyed; row-source errors are reported as
  `SOURCE_ERROR` with row details.
- Typed `ExcelIoError` codes: `INVALID_SPEC`, `IO_ERROR`, `ABORTED`,
  `WORKER_ERROR`, `ROW_VALIDATION`, `SOURCE_ERROR`.

**Packaging**

- ESM and CommonJS, readable and minified (`/min`) builds, worker bundles,
  source maps, and TypeScript declarations.

| Use case | Example |
| --- | --- |
| In-memory round-trip | [01-buffer-roundtrip](https://github.com/DipanshuHandoo/excelio/blob/main/examples/01-buffer-roundtrip.mjs) |
| Styled multi-sheet report to a file | [02-sales-report](https://github.com/DipanshuHandoo/excelio/blob/main/examples/02-sales-report.mjs) |
| Validated import with row errors | [03-import-validation](https://github.com/DipanshuHandoo/excelio/blob/main/examples/03-import-validation.mjs) |
| Stream a workbook to a destination | [04-streaming](https://github.com/DipanshuHandoo/excelio/blob/main/examples/04-streaming.mjs) |
| Workers and concurrent exports | [05-workers-and-concurrency](https://github.com/DipanshuHandoo/excelio/blob/main/examples/05-workers-and-concurrency.mjs) |
| Progress and cancellation | [06-progress-and-cancellation](https://github.com/DipanshuHandoo/excelio/blob/main/examples/06-progress-and-cancellation.mjs) |
| CommonJS and minified builds | [07-commonjs-and-minified](https://github.com/DipanshuHandoo/excelio/blob/main/examples/07-commonjs-and-minified.cjs) |
| XLSX to JSON | [08-excel-to-json](https://github.com/DipanshuHandoo/excelio/blob/main/examples/08-excel-to-json.mjs) |
| Service export migration | [09-service-exports](https://github.com/DipanshuHandoo/excelio/blob/main/examples/09-service-exports.mjs) |
| Stream rows from a paged source with rollover | [10-streaming-rows](https://github.com/DipanshuHandoo/excelio/blob/main/examples/10-streaming-rows.mjs) |

## Installation

```sh
npm install @dipanshuhandoo/excelio
```

Requires Node.js 22 or later. Node 22 and 24 LTS are the tested support matrix.
TypeScript consumers should install `typescript` and `@types/node` as development
dependencies; NodeNext resolution supports both import and require declarations.

## Quick Start

```js
import { excelio } from '@dipanshuhandoo/excelio';

const spec = [{
  workbook: 'sales',
  properties: { author: 'system', title: 'Sales Report' },
  sheets: [{
    sheet: 'Orders',
    columns: [
      { key: 'customer', header: 'Customer', width: 30 },
      { key: 'amount', type: 'number', format: '#,##0.00' },
      { key: 'date', type: 'date' }
    ],
    styling: { headerBold: true, freezeHeader: true, autoFilter: true },
    data: [{ customer: 'Alice', amount: 100, date: new Date('2026-01-15') }]
  }]
}];

const buffer = await excelio.write(spec);
await excelio.write(spec, { to: 'sales.xlsx' });
const result = await excelio.read(buffer);
console.log(result.data[0].sheets[0].data);
```

Stream a large export straight from a source, without building an array:

```js
async function* orders() {
  for await (const doc of Order.find({}).lean().cursor({ batchSize: 1000 })) {
    yield { customer: doc.customer, amount: doc.amount, date: doc.createdAt };
  }
}

await excelio.write([{
  workbook: 'orders',
  sheets: [{ sheet: 'Orders', columns: spec[0].sheets[0].columns, data: orders() }]
}], { to: 'orders.xlsx', signal, onProgress: console.log });
```

CommonJS exposes the same named exports:

```js
const { excelio, ExcelIoError } = require('@dipanshuhandoo/excelio');
```

Opt-in minified builds have the same API and include their own worker files:

```js
import { excelio } from '@dipanshuhandoo/excelio/min';
// CommonJS: const { excelio } = require('@dipanshuhandoo/excelio/min');
```

Minification reduces library code size, not ExcelJS's installed dependencies or
workbook-processing time. Both builds include source maps and declarations.

## Runnable Examples

See the [example guide](https://github.com/DipanshuHandoo/excelio/blob/main/examples/README.md)
for buffer round-trips, styled multi-sheet reports, validated imports, streaming,
workers/concurrent exports, progress/cancellation, CommonJS/minified usage,
XLSX-to-JSON conversion, service export migration, and streamed rows with sheet
rollover.
Examples live in the repository and use public package exports.

From a checkout, run all cases or select one:

```sh
npm run examples
npm run examples -- 03-import-validation
```

The runner builds first, checks results, and cleans up temporary spreadsheets.
Run the file examples directly after building to keep their XLSX/JSON output:

```sh
node examples/02-sales-report.mjs
node examples/04-streaming.mjs
node examples/08-excel-to-json.mjs
node examples/10-streaming-rows.mjs
```

These save to the ignored example-output/ directory by default.

## Options Reference

Common to `write`, `writeStream`, and `read`:

| Option | Default | Purpose |
| --- | --- | --- |
| `onProgress` | none | `(event) => void`, called every `batchSize` rows; writes also emit a final event |
| `batchSize` | 5000 | Progress interval; on write it must be a positive integer and also sets the backpressure/yield interval for streamed sheets |
| `signal` | none | `AbortSignal`; rejects with `ABORTED` |
| `logger` | none | `{ info?, warn?, error? }` |
| `useWorker` | `'auto'` | `'auto'`, `true`, or `false` |
| `workerThreshold` | `{ rows: 50_000, bytes: 10 MiB }` | Auto-mode offload thresholds |

Write only:

| Option | Default | Purpose |
| --- | --- | --- |
| `to` | none (returns Buffer/Buffer[]) | File path or writable stream |
| `beforeWrite` | none | `(row, ctx) => row`, synchronous transform |
| `maxRowsPerSheet` | 1,048,575 | Rows per worksheet before continuing on `"<name> (2)"` |
| `sharedStrings` | `'auto'` | `'auto'` (off when a workbook has a streamed sheet), `true`, or `false` |

Read only:

| Option | Default | Purpose |
| --- | --- | --- |
| `headerRow` | 1 | Row containing headers |
| `nullForBlank` | `true` | Empty cells become `null` |
| `trimStrings` | `true` | Trim string cells |
| `dateFormat` | `'iso'` | `'iso'`, `'date'`, or `'epoch'` |
| `afterRead` | none | `(row, ctx) => row`, synchronous transform |
| `validateRow` | none | `(row, ctx) => null \| { message, code?, column? }` |
| `onRowError` | `'collect'` | `'collect'` or `'fail'` |

Sheet spec: `sheet`, `columns`, `styling` (`headerBold`, `freezeHeader`,
`autoFilter`), `data` (array, iterable, or async iterable; optional), and
`totalRows` (progress hint for streamed sheets).

## API

### Writing

- `await excelio.write(spec)` returns a Buffer for one workbook or Buffer[] for
  multiple workbooks.
- `await excelio.write(spec, { to: filePath })` writes one workbook and returns
  the path. A writable stream target is also supported and returned unchanged.
- `excelio.writeStream(spec)` returns a readable stream synchronously, writes
  only the first workbook, and runs inline. Consume it or pipe it to a destination;
  handle asynchronous errors via the stream's error event or pipeline.
- A single `to` target cannot receive multiple workbooks. Omit it to receive
  separate buffers, or write each workbook individually. No ZIP wrapper is added.

`spec` is a non-empty array of workbooks, each containing a non-empty `sheets`
array. Sheets contain `sheet`, optional `columns` and `styling`, and optional
`data`. Workbook properties support author, company, title, created, and modified.
The `workbook` field is logical metadata, not an automatic filename.

`data` accepts an array, any iterable (for example a `function*` generator), or
any async iterable (an `async function*`, an object-mode stream, or any async
iterable such as a database cursor; excelio has no database dependency). Omitted
or null `data` writes an empty sheet. Iterables are consumed once, so pass a
separate iterable per sheet; reusing one iterable object across sheets throws
`INVALID_SPEC` (the same array may be reused).

```js
async function* rows() {
  for await (const doc of collection.find({}).batchSize(1000)) {
    yield { url: doc.url, status: doc.status };
  }
}

await excelio.write([{
  workbook: 'audit',
  sheets: [{ sheet: 'Pages', columns: [{ key: 'url' }, { key: 'status' }], data: rows() }]
}], { to: 'audit.xlsx' });
```

When excelio stops early (error or abort) it calls the iterator's `return()`,
which ends a `for await` loop in the generator and closes the underlying cursor.

Columns support key, header, width, type (`string`, `number`, `boolean`, `date`),
and Excel number-format string. Without explicit columns, keys are inferred from
the first row only (for iterables, the first row is peeked and still written).
Explicit columns are recommended for iterables: stable order and types. Width
defaults to 15. Sheet names replace forbidden Excel characters, truncate to 31
characters, and receive suffixes when duplicated. `logger.warn` reports truncation.

### Large Exports

- Combine iterable `data` with `to` (a file path or writable stream). Buffer
  output still collects the whole file in memory.
- Iterable sheets wait whenever about 8 MiB of produced XML is queued for
  compression or the target, checking every `batchSize` rows. Compression is
  slower than row production, so this matters for local files too.
- `sharedStrings: 'auto' | true | false` (default `'auto'`): the Excel
  shared-strings table is on unless a workbook has an iterable sheet. On gives a
  smaller file, but memory grows with the number of distinct strings (1,000,000
  rows of distinct URLs added 1.4–2.1 GB). Off keeps memory flat and writes
  plain strings as `t="str"` cells.
- `maxRowsPerSheet` (default and maximum 1,048,575, Excel's limit minus the
  header) continues a sheet on `"<name> (2)"`, `"<name> (3)"`, … with the same
  columns, header, and styling. It applies to arrays and iterables. Names are
  allocated after every logical sheet's own name, so a continuation never takes
  a name another sheet uses.
- `totalRows` on an iterable sheet is an optional hint used for progress
  `totalRows`; without it, progress `totalRows` is undefined.
- Measured with exceljs 4.4.0 on Node 22 (Windows): 1,000,000 rows with six
  distinct-string columns to a file grew RSS by about 50 MB and took 24–58 s
  across runs; a slow 64 KB sink and two 200,000-row sheets stayed under 25 MB. The
  bound relies on exceljs/archiver internals, so it is re-measured on every
  exceljs upgrade (`npm run test:memory`).

### Reading

```js
const result = await excelio.read('sales.xlsx');
// Also accepts a Buffer or readable stream.
```

The result contains:

```js
{
  ok: true,
  data: [{ workbook: undefined, sheets: [{ sheet: 'Orders', columns: [], data: [] }] }],
  errors: [],
  stats: { totalRows: 0, totalSheets: 1, durationMs: 0 }
}
```

Column entries contain `key` and `header`. Read data uses header text as object
keys, not the original write keys. A column written as `{ key: 'customer',
header: 'Customer' }` reads back as `{ Customer: 'Alice' }`. Keep header equal to
key for round-trippable field names. The original workbook label, properties,
column types, widths, and styling are not reconstructed in the read result.

Read options: `headerRow` (default 1), `nullForBlank` (true), `trimStrings` (true),
and `dateFormat` (`iso`, `date`, or `epoch`; default `iso`). Entirely blank rows
are skipped. Formula cells return cached results, not formula text; rich text
and hyperlinks return displayed text. Excel error values become null.
`totalRows` counts accepted rows after transformation and validation.

### Transforms and Validation

```js
const result = await excelio.read(buffer, {
  afterRead: row => ({ ...row, email: row.email?.toLowerCase() }),
  validateRow: row => row.amount < 0
    ? { code: 'NEGATIVE_AMOUNT', message: 'Amount must be non-negative', column: 'amount' }
    : null,
  onRowError: 'collect'
});

await excelio.write(spec, {
  beforeWrite: row => ({ ...row, amount: Number(row.amount) })
});
```

Hooks are synchronous. They receive `(row, ctx)` where ctx contains workbookIndex,
workbookName, sheetIndex, sheetName, and the 1-based Excel rowNumber. On write,
ctx also has `sheetPart` (1, 2, … after `maxRowsPerSheet` rollover); sheetIndex
stays the spec index, sheetName is the actual worksheet name, and rowNumber
restarts at 2 on each continuation sheet. Read applies
afterRead before validateRow. Invalid rows are excluded from data and collected
in errors by default; `onRowError: 'fail'` throws on the first issue. A hook
returning null/undefined keeps the original row. Write transforms do not add
columns beyond the explicit or inferred column set.

### Workers, Progress, and Cancellation

```js
await excelio.write(spec, { useWorker: true });
await excelio.read(buffer, { useWorker: false });
await excelio.write(spec, { workerThreshold: { rows: 100_000, bytes: 50 * 1024 * 1024 } });
```

`useWorker` defaults to `auto`. Writes offload at 50,000 total rows; reads offload
at 10 MiB for buffers or file paths. Threshold comparisons are inclusive.
Each offloaded operation creates one worker, then terminates it. There is no
reusable worker pool, concurrency cap, or parallel splitting of a workbook.
Submit independent operations concurrently only when memory and CPU permit.

Functions (`beforeWrite`, `afterRead`, `validateRow`, `onProgress`), `logger`,
`signal`, stream inputs, writable stream destinations, and iterable sheet data
force inline execution, even when useWorker is true (with iterable data,
`logger.warn` reports the ignored useWorker). Worker payloads must be structured-cloneable;
functions or unsupported values embedded in row data cannot be sent to workers.
Bundlers/serverless deployments must preserve the emitted entrypoint and its
sibling workers directory. Import the package normally rather than copying a
single bundled file.

`onProgress` receives phase and rowsProcessed plus available workbook/sheet
metadata; write events include sheetIndex, sheetName, and sheetPart. `batchSize`
defaults to 5,000 and must be a positive integer. Writes emit a final
notification; reads notify at batch boundaries.

`signal` is checked between rows and sheets. Iterable sheets check it on every
row and yield to the event loop every `batchSize` rows, so an abort fired from a
timer or request handler is observed during long iterable writes. A source stuck
inside `next()` cannot be interrupted by excelio; pass the same signal to the
source. Array sheets still run as a synchronous loop, and initial XLSX loading
and commit cannot be interrupted.

### Values and Errors

Date columns accept Date, ISO date strings, and epoch milliseconds. Number
columns convert finite numeric strings; boolean columns recognize true/false,
the strings 'true'/'false', and 1/0. Unsupported coercions are left unchanged;
these hints are not schema validation. Multiline and multilingual text is supported.

`ExcelIoError` exposes code, optional cause, and optional details. Codes include
INVALID_SPEC, IO_ERROR, ABORTED, WORKER_ERROR, ROW_VALIDATION, and SOURCE_ERROR.
SOURCE_ERROR means an iterable sheet's source threw (or returned a non-object
iterator result); `cause` is the original error and `details` holds
workbookIndex, sheetIndex, sheetName, and rowsRead. If the signal is already
aborted when the source fails, the error is ABORTED with the source error as
`cause`. Collected read
errors contain workbook index, sheet, row, column, code, and message. `ok` is
false when collected errors exist; callers decide whether to accept partial data.
Some lower-level ExcelJS/write errors can propagate as ordinary Error instances.

On a failed write, a file-path target is closed and removed; a caller's writable
stream is unpiped but neither ended nor destroyed; the original error is thrown,
with `err.cleanupError` set if cleanup itself failed. On success a caller's
stream is ended by the write, as before.

## Memory and Scope

Writes use ExcelJS's streaming XLSX writer and commit rows incrementally. Array
`data` holds all rows in memory, and so does the shared-strings table when it is
on. Iterable `data` with the default `sharedStrings` and a file or stream target
keeps write memory bounded (see Large Exports). Buffer output collects the
complete file either way; direct file/stream output avoids that final buffer.
Reading currently loads the entire workbook and collects accepted rows.
Workers keep processing off the main thread but do not guarantee lower total
memory or faster completion; cloning large inputs adds overhead.

Browsers, legacy XLS, charts, pivots, macros, encryption, and advanced formatting
are outside this library's API. Only features described above are supported.

## Development and Releases

```sh
npm ci
npm run verify
npm run build:min
npm run test:perf
npm run pack:check
npm run publish:dry-run
```

`build` produces dist/esm, dist/cjs, dist/esm-min, and dist/cjs-min with bundled
worker entrypoints and matching declarations. `build:min` regenerates only the
minified outputs; run a full build for a release. The only direct runtime
dependency is ExcelJS; esbuild, TypeScript, and Node types are development tools.

See [Contributing](https://github.com/DipanshuHandoo/excelio/blob/main/CONTRIBUTING.md),
[Releasing](https://github.com/DipanshuHandoo/excelio/blob/main/RELEASING.md),
[Security](https://github.com/DipanshuHandoo/excelio/blob/main/SECURITY.md), and
[CHANGELOG.md](CHANGELOG.md). Public npm availability
starts with the first approved publication; repository setup does not publish it.

## License

MIT. See [LICENSE](LICENSE).