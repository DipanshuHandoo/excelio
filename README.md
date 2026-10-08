# ExcelIO

`@dipanshuhandoo/excelio` converts JavaScript row objects to XLSX workbooks and
reads XLSX files into structured data. It supports streaming writes, multiple
sheets/workbooks, row transforms and validation, and configurable worker-thread
offloading. It is a Node.js library, not a browser library.

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
array. Sheets contain `sheet`, optional `columns` and `styling`, and a `data`
array. Workbook properties support author, company, title, created, and modified.
The `workbook` field is logical metadata, not an automatic filename.

Columns support key, header, width, type (`string`, `number`, `boolean`, `date`),
and Excel number-format string. Without explicit columns, keys are inferred from
the first row only. Width defaults to 15. Sheet names replace forbidden Excel
characters, truncate to 31 characters, and receive suffixes when duplicated.
`logger.warn` reports truncation.

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
workbookName, sheetIndex, sheetName, and the 1-based Excel rowNumber. Read applies
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
`signal`, stream inputs, and writable stream destinations force inline execution,
even when useWorker is true. Worker payloads must be structured-cloneable;
functions or unsupported values embedded in row data cannot be sent to workers.
Bundlers/serverless deployments must preserve the emitted entrypoint and its
sibling workers directory. Import the package normally rather than copying a
single bundled file.

`onProgress` receives phase and rowsProcessed plus available workbook/sheet
metadata. `batchSize` defaults to 5,000 and controls notifications, not memory
or yielding. Writes emit a final notification; reads notify at batch boundaries.
`signal` is checked between rows and sheets, but cannot interrupt initial XLSX
loading, commit, or synchronous row loops. Do not rely on timely cancellation of
an in-progress operation from an event-loop callback.

### Values and Errors

Date columns accept Date, ISO date strings, and epoch milliseconds. Number
columns convert finite numeric strings; boolean columns recognize true/false,
the strings 'true'/'false', and 1/0. Unsupported coercions are left unchanged;
these hints are not schema validation. Multiline and multilingual text is supported.

`ExcelIoError` exposes code, optional cause, and optional details. Codes include
INVALID_SPEC, IO_ERROR, ABORTED, WORKER_ERROR, and ROW_VALIDATION. Collected read
errors contain workbook index, sheet, row, column, code, and message. `ok` is
false when collected errors exist; callers decide whether to accept partial data.
Some lower-level ExcelJS/write errors can propagate as ordinary Error instances.

## Memory and Scope

Writes use ExcelJS's streaming XLSX writer and commit rows incrementally, but
input arrays and shared strings remain in memory; buffer output also collects
the complete file. Direct file/stream output avoids that final buffer.
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