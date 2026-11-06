# ExcelIO 1.1.0 Plan: Streaming Row Sources, Sheet Rollover, Shared-Strings Control

**Status:** implemented on branch `streaming` (steps 0–11; release §9 pending owner approval). Open questions use their defaults; Q5 budget = 65 MB (scenario A measured 50.3–51 MB). Q2 manual check passed: a `sharedStrings: false` workbook opened without a repair prompt (2026-11-06). Target release: **1.1.0** (minor, backward compatible for JavaScript callers; see Q9).

**Audience:** an implementer who has not seen the discussion that produced this plan.
- Every decision in §2 is final unless it appears in **§11 Open questions**, which also gives the default to use when no answer is available.
- §1 facts were verified against `src/` at 1.0.0, exceljs 4.4.0, archiver 5.3.2, zip-stream 4.1.1 and compress-commons, as installed in `node_modules`.
- The plan was reviewed independently once; that review's findings are already folded in.

---

## 0. Why this exists, and what it is not

Today every sheet's `data` must be an in-memory array. To export 600,000 rows, a caller first builds a 600,000-element array, so memory grows with the export size. The XLSX writer inside excelio is already exceljs's *streaming* writer; only the input side is not streaming, and (fact 4) the shared-strings table is not either.

- The first consumer is the Spidy Web Excel report in `gwp-spidy-v3`: up to about 9 sheets, the largest ("Links") 600k to 2M rows, read from MongoDB.
- The capability is **generic**. No MongoDB or other database support is added to excelio, and no runtime dependency is added. Excelio accepts the JavaScript **iteration protocols** (`Symbol.iterator`, `Symbol.asyncIterator`) and only ever calls `next()` and `return()`.

| Source | How a caller passes it |
|---|---|
| MongoDB driver `FindCursor` / `AggregationCursor` | the cursor itself (it is async-iterable) |
| Mongoose | `Model.find(q).lean().cursor()` |
| Node `Readable` in object mode (`pg-query-stream`, `csv-parse`, `readline`) | the stream itself (streams are async-iterable) |
| Paginated REST API | an `async function*` that fetches page after page |
| Computed rows | a `function*` generator, or an `async function*` mapping another source |

**Recommended caller pattern (Spidy Web).** Wrap the cursor in an `async function*` with `for await`. When excelio stops early (error or abort) it calls the generator's `return()`, which ends the `for await` and closes the cursor.

```js
async function* linkRows(runId) {
  for await (const e of SpidyWebEdge.find({ runId }).sort({ _id: 1 }).lean().cursor({ batchSize: 1000 })) {
    yield { source: e.sourceUrl, target: e.normalizedUrl, occurrences: e.occurrences /* … */ };
  }
}

await excelio.write(
  [{ workbook: 'report', sheets: [{ sheet: 'Links', columns, data: linkRows(runId) }] }],
  { to: tmpFile, signal, onProgress, maxRowsPerSheet: 1_000_000 }
);
```

---

## 1. Facts about the current code (verified)

**excelio 1.0.0 (`src/`)**

1. `writer.js#writeOneWorkbook` loops `for (let i = 0; i < data.length; i++)` with `data = sheetSpec.data || []`. A non-array `data` (for example a generator) silently writes **zero rows**, because `.length` is undefined. This is a latent bug.
2. `totalRows = sheets.reduce((n, s) => n + (s.data?.length || 0), 0)`.
3. `columns.js#resolveColumns(declared, data)` infers keys from `data[0]` when `columns` is omitted.
4. The WorkbookWriter is created with `useSharedStrings: true`. exceljs keeps the **entire shared-strings table in memory** (`new SharedStrings()` in `workbook-writer.js`) and writes it only at `commit()`. With mostly distinct strings (URLs, ids, emails), memory grows with the row count even if the input streams.
5. `index.js#write` calls `maybeUseWorker()` **before** any validation. `validateSpec` runs inside `writeSpec`, so for worker-routed calls it runs inside the worker.
6. `workerPool.js#decideMode` forces inline for read streams, hooks (`beforeWrite`, `afterRead`, `validateRow`, `onProgress`), `logger`, `signal`, and writable `to`. `estimateRows` sums `sheet.data?.length`. `sanitiseOptions` strips functions/signal/logger/stream `to` and passes everything else (so new plain options such as `maxRowsPerSheet` and `sharedStrings` would reach a worker).
7. `signal` is checked only between rows of a **synchronous** loop, so an abort fired from an event-loop callback cannot be observed mid-sheet (the README says so).
8. On a write error, a file-path target's `WriteStream` stays open and a partial `.xlsx` remains on disk. The archiver also keeps running.
9. `batchSize` is not validated. A value of `0` makes `rowsProcessed % 0` `NaN`, so progress events silently stop.
10. Existing worker tests assert routing by outcome, for example "forced worker writes with hooks fall back inline". They do not spy on `worker_threads.Worker`.
11. `package.json` is `1.0.0` (published), but `CHANGELOG.md` still says `## 0.1.0 (Unreleased)` (Q11).
12. Sheet names come from `sanitiseSheetName(raw, usedNames, logger)`: it replaces forbidden characters, truncates to 31 characters, and on a duplicate appends ` (2)`, ` (3)` and so on, truncating the base so the result fits in 31 characters.

**exceljs 4.4.0 / archiver 5.3.2 (the output pipeline)**

13. Byte path for one worksheet:
    `worksheet.stream` (a `StreamBuf`, created by `workbook-writer.js#_openStream` with `{ bufSize: 65536, batch: true }`)
    → `.pipe()` into a per-entry `DeflateCRC32Stream` created by compress-commons `ZipArchiveOutputStream#_smartStream`
    → `process.pipe(zipOutputStream, { end: false })`
    → archiver (`wb.zip`, `this._module.pipe(this)`)
    → `wb.zip.pipe(sink)`.
14. **`StreamBuf.write()` always returns `true`.** In batch mode it pushes each full 64 KB buffer into its pipes with `this._pipe(...)`, **not awaited**, so the `DeflateCRC32Stream`'s writable buffer absorbs everything that cannot be compressed yet. Every later hop uses `pipe()` and does respect backpressure. Excelio therefore cannot rely on `write()` return values. It must measure the backlog itself (D7).
15. The backlog lives in these places, all of them **internal, undocumented fields**:
    - `ws.stream.buffers` (array of `ReadWriteBuf`, each with a `length` getter): bytes not yet handed to a pipe;
    - `ws.stream.pipes[0]`: **not** the `DeflateCRC32Stream`. `StreamBuf` never initialises `_readableState`, so archiver's `archiver-utils.normalizeInputSource` wraps it in a `PassThrough` from **readable-stream 2.3.8** at `append()` time, and `StreamBuf` pipes into that wrapper from the start. readable-stream 2.x has **no `writableLength`/`readableLength` getters**, so the measure must read `_writableState.length` and `_readableState.length`. This wrapper holds the whole backlog of a worksheet whose entry is waiting in archiver's queue and of the active entry while deflate lags (verified in S0: it was about 100 % of the measured backlog);
    - `wb.zip._module` (the `ZipArchiveOutputStream`) and `wb.zip` (archiver, readable-stream 3): `writableLength` and `readableLength`;
    - the sink: `writableLength`.
    
    `wb._worksheets` is the array of WorksheetWriters (index 0 empty).

    **S0 result (Node 22.23, Windows, 6 columns of distinct strings, `useSharedStrings: false`):**

    | Scenario | Sink | No waiting: peak RSS growth / backlog | With `waitForCapacity`: peak RSS growth / backlog | Total time (no wait → wait) |
    |---|---|---|---|---|
    | B, 300k rows, 1 sheet | 64 KB, `setTimeout(2)` per write | 141 MB / 102 MB | 61 MB / 9.8 MB | 24.8 s → 24.7 s |
    | C, 2 × 200k rows | 64 KB, `setTimeout(2)` | 186 MB / 135 MB | 47 MB / 9.8 MB | 33.2 s → 32.1 s |
    | C | synchronous callback | 176 MB / 137 MB | 59 MB / 9.8 MB | 10.1 s → 9.3 s |

    Deflate is slower than row production, so the backlog builds up **even with an instant sink**. D7 is needed for file targets too, not only slow sinks. Waiting does not slow the overall write. Throughput was about 43k rows/s with an instant sink.
16. Archiver processes entries one at a time (`async.queue(…, 1)`), in the order worksheets were created.
17. With shared strings **off**, plain string cells are written as `t="str"` with a `<v>` value (`xlsx/xform/sheet/cell-xform.js`). `inlineStr` is used only for rich text.
18. Setting `ws.columns` with headers writes the header row on a streaming worksheet, so each new worksheet gets its header.
19. exceljs attaches its own `error` listeners to the sink and archiver **only in `_finalize()`**, at commit. An early sink error, such as a missing directory, is not seen until then, and `commit()` may wait for a `finish` that never comes.
20. Archiver `abort()` sets `_state.aborted`, kills its queues and shuts down when idle. It does not drain or end an active entry source.
21. exceljs already depends on `jszip` (`^3.10.1`) and `unzipper`. Neither is a direct dependency of excelio.

---

## 2. Decisions

- **D1. Accepted `data` values per sheet:**
  - `Array` (unchanged);
  - any **sync iterable** (generator, `Set`, custom iterable) except strings;
  - any **async iterable** (async generator, cursor, object-mode stream);
  - `undefined`/`null`, which means an empty sheet (unchanged).
  
  Anything else (plain object, number, string, boolean) → `ExcelIoError` `INVALID_SPEC`: `Sheet "<name>": data must be an array or an iterable of row objects`.
- **D2. Arrays keep today's synchronous fast path.** There is no per-row `await` and output is unchanged. Only iterable sheets use the new async path. Both paths write rows through one shared function (§4.4), so formatting, coercion and `beforeWrite` behave identically.
- **D3. Columns for iterable sheets.** Use `columns` when given; otherwise **peek the first row** and infer keys from it (as arrays do from `data[0]`). The peeked row is written as the first data row; nothing is lost. An empty iterable without `columns` gives an empty sheet with no header, the same as an empty array today. The README recommends explicit `columns` (stable order, types); they are not required.
- **D4. Validation runs before any routing.** A new `preflightWrite(spec, options)` runs in `index.js` before `maybeUseWorker()` for `write`, and at the start of `writeStream`. `writeSpec` still validates, as defence inside the worker.
- **D5. Iterable sheets always run inline.** A spec containing any iterable sheet is written in the calling thread, even with `useWorker: true` (iterators cannot be structured-cloned). `logger.warn` is called once when `useWorker === true` was overridden.
- **D6. New write option `sharedStrings: 'auto' | true | false`, default `'auto'` (Q1).**
  - `'auto'` means *enabled unless any sheet of the workbook has iterable data*, so array-only and empty workbooks are unchanged.
  - `false` writes plain strings as `t="str"` (fact 17): the file is somewhat larger and memory stays flat (Q2).
  - `true` and `false` force the mode. The setting applies per workbook.
- **D7. Backlog-bounded writing (iterable path only).**
  - Every `batchSize` rows (default 5,000), the writer calls `waitForCapacity()` (§4.4). It polls a backlog measure (fact 15) until the measure is at or below `OUTPUT_HIGH_WATER_MARK` (8 MiB, an internal constant; Q6), then yields one `setImmediate` tick.
  - While it waits it checks the `signal` (throws `ABORTED`) and the sink failure state (throws the sink error), so it can never wait forever on a dead stream.
  - The measure reads exceljs/archiver internals. Every read is optional-chained and defaults to `0`, so an incompatible exceljs version degrades to "yield only" instead of crashing.
  - Spike **S0** (§8) must confirm the measure before the rest is built. The array path is unchanged.
- **D8. Rollover with `maxRowsPerSheet`**, default and maximum `1_048_575` (Excel's 1,048,576 rows minus the header). It must be an integer `1..1_048_575`, otherwise `INVALID_SPEC`.
  - When a worksheet reaches the limit, it is committed and a continuation worksheet is opened with the same columns, styling and header.
  - It applies to arrays and iterables (Q3).
  - **Naming:** before writing, the first-part name of every logical sheet in the workbook is reserved, using today's `sanitiseSheetName` in today's order, so array specs keep exactly today's names. Continuation names are then allocated with `sanitiseSheetName(baseName, used)`, where `used` holds every reserved and created name. Example: logical sheets `Data` and `Data (2)`, and `Data` overflows → its continuation is `Data (3)`, and the second logical sheet keeps `Data (2)`.
- **D9. Errors.**
  - **Source errors:** a throw from obtaining the iterator (`[Symbol.(async)iterator]()`), from `next()`, or a `next()` result that is not an object becomes `ExcelIoError` code **`SOURCE_ERROR`**, with `cause` = the original error and `details: { workbookIndex, sheetIndex, sheetName, rowsRead }`.
  - **Abort precedence:** if `signal.aborted` is true when the source throws, the error is `ABORTED` with the source error as `cause`, not `SOURCE_ERROR`.
  - `beforeWrite` errors keep their current behaviour (rethrown as is).
  - On any error or abort while an iterable sheet is open, the iterator's `return()` is called (best effort, errors ignored).
- **D10. Failure cleanup**, the same for array and iterable sheets (Q4). On any error after output started, run this abort sequence, then rethrow the **original** error:
  1. Unpipe archiver from the sink (`wb.zip.unpipe(sink.writable)`), so a caller's stream is not ended.
  2. Attach no-op `error` listeners to `wb.zip`, `wb.zip._module` and every open worksheet's `stream.pipes[*]`, then `wb.zip.abort()`. Late errors from the half-built archive must not become unhandled.
  3. Destroy every library-owned target: the Buffer collector (`writeOneWorkbookToBuffer`) and the file `WriteStream` (`writeStream()` already destroys its PassThrough).
  4. **File targets:** wait for the `WriteStream` `close` event, then `rm(target, { force: true })`. Waiting for `close` matters on Windows, where `rm` fails with `EPERM` while the handle is open.
  5. If cleanup itself fails, attach it to the original error as `err.cleanupError`. Never replace the original error.
  
  **Caller-supplied writable targets:** on success they are ended by the archiver pipe (today's behaviour, unchanged). On failure excelio unpipes and **does not** end or destroy them. Document both.
- **D11. Early sink failures.** `resolveSink` attaches `error`/`close` listeners immediately and exposes `sink.failure` (the first error) and `sink.failed` (a promise rejecting on the first error, with a no-op catch attached so it is never unhandled).
  - The iterable path checks `sink.failure` in `waitForCapacity`.
  - `commit` is awaited as `Promise.race([wb.commit(), sink.failed])`, so an error such as a missing directory rejects promptly instead of hanging (fact 19).
- **D12. Progress.**
  - `totalRows` is the sum of array lengths plus an optional new `sheet.totalRows` hint for iterable sheets. If any iterable sheet has no hint, `totalRows` is `undefined`.
  - Every write progress event, including the final one, carries `sheetIndex` (logical, the spec index; Q7), `sheetName` (the actual worksheet name, including any ` (2)` suffix) and `sheetPart` (1, 2, …).
  - The final event uses the last worksheet opened. A workbook whose last logical sheet is empty reports that sheet's first-part name and `sheetPart: 1`.
  - `RowContext` (for `beforeWrite`) gains `sheetPart`. `sheetIndex` stays logical. `sheetName` is the actual worksheet name. `rowNumber` is the Excel row number, so it restarts at 2 in each continuation.
- **D13. `batchSize` is validated** as a positive integer (`>= 1`), for every write (Q12). It now bounds memory and abort latency, not just progress frequency.
- **D14. Reusing one iterable object across sheets is rejected** (Q8). It is a conservative rule: generators, cursors and streams are single-use, and a reused one silently yields an empty second sheet. Excelio cannot tell a reusable collection (a `Set`) from a single-use iterator, so it rejects any **non-array** iterable object referenced by more than one sheet, with `INVALID_SPEC` `Sheet "<b>": its data source is already used by sheet "<a>". Pass a separate iterable (or an array) per sheet.` The same **array** may be reused (unchanged).
- **D15. Out of scope for 1.1.0:**
  - streaming reads (`read()` still loads the whole workbook);
  - database adapters;
  - parallel or chunked writing;
  - zipping several workbooks;
  - an async `beforeWrite`;
  - a public option for the high-water mark.

---

## 3. Public API after the change (`src/excelio.d.ts`)

```ts
export type ExcelIoErrorCode =
  | 'INVALID_SPEC' | 'IO_ERROR' | 'ABORTED' | 'WORKER_ERROR' | 'ROW_VALIDATION' | 'SOURCE_ERROR';

/** Rows for one sheet: an array, or any sync/async iterable (generator, database cursor, object-mode stream).
 *  Iterables are consumed once and always written inline. */
export type RowSource<Row> = readonly Row[] | Iterable<Row> | AsyncIterable<Row>;

export interface SheetSpec<Row = Record<string, unknown>> {
  sheet: string;
  columns?: ColumnSpec[];
  styling?: SheetStyling;
  /** Omitted or null means an empty sheet. */
  data?: RowSource<Row> | null;
  /** Optional row-count hint used for progress `totalRows` when `data` is an iterable. */
  totalRows?: number;
}

export interface WriteOptions<Row = Record<string, unknown>> extends CommonOptions {
  to?: string | Writable;
  beforeWrite?: (row: Row, ctx: RowContext) => Row | void;
  /** Data rows per worksheet before continuing on "<name> (2)". Default and maximum 1,048,575. */
  maxRowsPerSheet?: number;
  /** Excel shared-strings table. 'auto' (default): on unless any sheet has iterable data. */
  sharedStrings?: 'auto' | boolean;
}

export interface RowContext { /* existing fields */ sheetPart?: number }
export interface ProgressEvent { /* existing fields */ sheetPart?: number }
```

`data` becomes optional in the types, matching the runtime (D1). That is a widening, so it is non-breaking. Everything else is unchanged.

---

## 4. Implementation, file by file

### 4.1 New `src/rowSource.js`

```js
import { ExcelIoError } from './errors.js';

/** 'empty' | 'array' | 'sync' | 'async' | 'invalid' */
export function rowSourceKind(data) {
  if (data === undefined || data === null) return 'empty';
  if (Array.isArray(data)) return 'array';
  if (typeof data !== 'object' && typeof data !== 'function') return 'invalid';   // strings, numbers, booleans
  if (typeof data[Symbol.asyncIterator] === 'function') return 'async';
  if (typeof data[Symbol.iterator] === 'function') return 'sync';
  return 'invalid';
}

export const isIterableSource = (data) => {
  const kind = rowSourceKind(data);
  return kind === 'sync' || kind === 'async';
};

export const specHasIterable = (spec) =>
  Array.isArray(spec) && spec.some((wb) => (wb?.sheets || []).some((s) => isIterableSource(s?.data)));

/** Closes a partially consumed iterator (cursor, stream). Never throws. */
export async function closeIterator(iterator) {
  if (typeof iterator?.return !== 'function') return;
  try { await iterator.return(); } catch { /* the source is already failing or closed */ }
}

/** SOURCE_ERROR, or ABORTED when the caller's signal caused the failure (D9). */
export function sourceFailure(err, details, signal) {
  if (signal?.aborted) return new ExcelIoError('Aborted', { code: 'ABORTED', cause: err, details });
  return new ExcelIoError(`Row source failed for sheet "${details.sheetName}": ${err?.message ?? err}`, {
    code: 'SOURCE_ERROR', cause: err, details
  });
}
```

### 4.2 Validation (`src/writer.js`, exported; used by `index.js` and `writeSpec`)

```js
export const EXCEL_MAX_DATA_ROWS = 1_048_575;

/** D4: called before routing (index.js) and again inside writeSpec/writeSpecStream. */
export function preflightWrite(spec, options = {}) {
  validateSpec(spec);
  validateWriteOptions(options);
}

function validateSpec(spec) {
  // existing checks: non-empty array of workbook objects, each with a non-empty sheets array
  const owners = new Map();                                    // D14: iterable object → first sheet name
  spec.forEach((wb, wbIndex) => {
    wb.sheets.forEach((s, i) => {
      if (!s || typeof s !== 'object') {
        throw new ExcelIoError(`Workbook ${wbIndex}: sheet entries must be objects`, { code: 'INVALID_SPEC' });
      }
      const name = s.sheet || `Sheet ${i + 1}`;
      const kind = rowSourceKind(s.data);
      if (kind === 'invalid') {
        throw new ExcelIoError(`Sheet "${name}": data must be an array or an iterable of row objects`, { code: 'INVALID_SPEC' });
      }
      if (kind === 'sync' || kind === 'async') {
        if (owners.has(s.data)) {
          throw new ExcelIoError(
            `Sheet "${name}": its data source is already used by sheet "${owners.get(s.data)}". Pass a separate iterable (or an array) per sheet.`,
            { code: 'INVALID_SPEC' });
        }
        owners.set(s.data, name);
      }
      if (s.totalRows !== undefined && !(Number.isInteger(s.totalRows) && s.totalRows >= 0)) {
        throw new ExcelIoError(`Sheet "${name}": totalRows must be a non-negative integer`, { code: 'INVALID_SPEC' });
      }
    });
  });
}

function validateWriteOptions({ maxRowsPerSheet, sharedStrings, batchSize } = {}) {
  if (maxRowsPerSheet !== undefined &&
      !(Number.isInteger(maxRowsPerSheet) && maxRowsPerSheet >= 1 && maxRowsPerSheet <= EXCEL_MAX_DATA_ROWS)) {
    throw new ExcelIoError(`maxRowsPerSheet must be an integer between 1 and ${EXCEL_MAX_DATA_ROWS}`, { code: 'INVALID_SPEC' });
  }
  if (sharedStrings !== undefined && sharedStrings !== 'auto' && typeof sharedStrings !== 'boolean') {
    throw new ExcelIoError(`sharedStrings must be 'auto', true or false`, { code: 'INVALID_SPEC' });
  }
  if (batchSize !== undefined && !(Number.isInteger(batchSize) && batchSize >= 1)) {
    throw new ExcelIoError(`batchSize must be a positive integer`, { code: 'INVALID_SPEC' });   // D13
  }
}
```

### 4.3 `src/index.js` and `src/workerPool.js`

```js
// index.js
async write(spec, options = {}) {
  preflightWrite(spec, options);                                           // D4
  return maybeUseWorker('write', spec, options, () => writeSpec(spec, options));
},
writeStream(spec, options = {}) {
  preflightWrite(spec, options);
  return writeSpecStream(spec, options);
},
```

```js
// workerPool.js
import { specHasIterable } from './rowSource.js';

function decideMode(op, payload, options) {
  const useWorker = options.useWorker ?? 'auto';
  if (op === 'read' && payload && typeof payload.pipe === 'function') return { mode: 'inline', reason: 'stream input' };
  if (op === 'write' && specHasIterable(payload)) {                        // D5
    if (useWorker === true) options.logger?.warn?.('excelio: iterable sheet data is written inline; useWorker ignored');
    return { mode: 'inline', reason: 'iterable sheet data' };
  }
  // …unchanged…
}

// estimateRows: count arrays only — Array.isArray(sheet.data) ? sheet.data.length : 0
```

Note: `logger` presence already forces inline today (fact 6), so the warning fires in the inline path; that is intended.

### 4.4 `src/writer.js`: `writeOneWorkbook` restructured

Keep the signature `writeOneWorkbook(workbookSpec, wbIndex, target, options)`. `writeSpec` and `writeSpecStream` call `preflightWrite` at their start (in place of `validateSpec`).

```js
import { rm } from 'node:fs/promises';
import { setImmediate as tick, setTimeout as sleep } from 'node:timers/promises';
import { rowSourceKind, isIterableSource, closeIterator, sourceFailure } from './rowSource.js';

const OUTPUT_HIGH_WATER_MARK = 8 * 1024 * 1024;   // D7, internal (Q6)

async function writeOneWorkbook(workbookSpec, wbIndex, target, options, sinkOptions = {}) {
  const { signal, onProgress, beforeWrite, batchSize = DEFAULT_BATCH, logger } = options;
  const maxRowsPerSheet = options.maxRowsPerSheet ?? EXCEL_MAX_DATA_ROWS;
  const sheets = workbookSpec.sheets || [];
  const sink = resolveSink(target, sinkOptions); // D11: failure tracking starts now
  let wb = null;                               // declared outside try so the catch can abort it (D10)

  try {
    wb = new ExcelJS.stream.xlsx.WorkbookWriter({
      stream: sink.writable,
      useStyles: true,
      useSharedStrings: resolveSharedStrings(options.sharedStrings, sheets)   // D6
    });
    applyWorkbookProperties(wb, workbookSpec.properties);

    const used = new Set();
    const firstNames = sheets.map((s, i) => sanitiseSheetName(s.sheet || `Sheet ${i + 1}`, used, logger)); // D8 reserve
    const totalRows = computeTotalRows(sheets);
    const state = { rowsProcessed: 0, last: { sheetIndex: 0, sheetName: firstNames[0], sheetPart: 1 } };
    const emit = () => onProgress?.({ phase: 'write', workbookIndex: wbIndex, ...state.last, rowsProcessed: state.rowsProcessed, totalRows });

    for (let sIndex = 0; sIndex < sheets.length; sIndex++) {
      throwIfAborted(signal);
      const sheetSpec = sheets[sIndex];
      const writer = createSheetWriter({
        wb, workbookSpec, wbIndex, sIndex, sheetSpec, firstName: firstNames[sIndex], used, logger, beforeWrite, maxRowsPerSheet,
        onOpen: (sheetName, sheetPart) => { state.last = { sheetIndex: sIndex, sheetName, sheetPart }; },
        onRow: () => {
          state.rowsProcessed++;
          if (onProgress && state.rowsProcessed % batchSize === 0) emit();
        }
      });

      const kind = rowSourceKind(sheetSpec.data);
      if (kind === 'sync' || kind === 'async') {
        await writeIterableSheet({ writer, wb, sink, sheetSpec, kind, signal, batchSize, wbIndex, sIndex });
      } else {
        const data = sheetSpec.data || [];
        writer.start(resolveColumns(sheetSpec.columns, data));
        for (let i = 0; i < data.length; i++) {
          throwIfAborted(signal);
          writer.write(data[i]);
        }
      }
      writer.finish();
    }

    await Promise.race([wb.commit(), sink.failed]);                       // D11
    if (sink.cleanup) await Promise.race([sink.cleanup(), sink.failed]);
    emit();                                                                // final event (D12)
  } catch (err) {
    await abortWorkbook(wb, sink, err);                                    // D10
    throw err;
  }
}
```

The sheet writer handles rollover (D8) and is the single row-writing path for arrays and iterables (D2):

```js
function createSheetWriter({ wb, workbookSpec, wbIndex, sIndex, sheetSpec, firstName, used, logger, beforeWrite, maxRowsPerSheet, onOpen, onRow }) {
  const baseName = sheetSpec.sheet || `Sheet ${sIndex + 1}`;
  let cols = [];
  let ws = null;
  let sheetName = null;
  let part = 0;
  let rowsInPart = 0;

  const open = () => {
    part += 1;
    rowsInPart = 0;
    sheetName = part === 1 ? firstName : sanitiseSheetName(baseName, used, logger);   // continuation: "<base> (n)"
    const wsOptions = {};
    if (sheetSpec.styling?.freezeHeader) wsOptions.views = [{ state: 'frozen', ySplit: 1 }];
    ws = wb.addWorksheet(sheetName, wsOptions);
    if (cols.length) {
      ws.columns = cols.map((c) => ({ key: c.key, header: c.header, width: c.width, style: c.format ? { numFmt: c.format } : undefined }));
    }
    applyStylingPostColumns(ws, sheetSpec.styling, cols.length);
    onOpen(sheetName, part);
  };

  return {
    get sheetName() { return sheetName ?? firstName; },
    start(resolvedColumns) { cols = resolvedColumns; open(); },
    write(row) {
      if (rowsInPart >= maxRowsPerSheet) { ws.commit(); open(); }
      let out = row;
      if (beforeWrite) {
        out = beforeWrite(row, {
          workbookIndex: wbIndex, workbookName: workbookSpec.workbook,
          sheetIndex: sIndex, sheetName, sheetPart: part, rowNumber: rowsInPart + 2
        }) ?? row;
      }
      const values = cols.length
        ? cols.reduce((acc, c) => { acc[c.key] = valueForWrite(out[c.key], c.type); return acc; }, {})
        : out;
      ws.addRow(values).commit();
      rowsInPart++;
      onRow();
    },
    finish() { ws?.commit(); }
  };
}
```

`RowContext.sheetPart` is new. Today's fields and their values are unchanged for non-rollover sheets (`rowNumber` = index + 2).

The iterable path:

```js
async function writeIterableSheet({ writer, wb, sink, sheetSpec, kind, signal, batchSize, wbIndex, sIndex }) {
  let rowsRead = 0;
  const details = () => ({ workbookIndex: wbIndex, sheetIndex: sIndex, sheetName: writer.sheetName, rowsRead });

  let iterator;
  try {
    iterator = kind === 'async' ? sheetSpec.data[Symbol.asyncIterator]() : sheetSpec.data[Symbol.iterator]();
  } catch (err) {
    throw sourceFailure(err, details(), signal);                           // D9: acquisition failure
  }

  const next = async () => {
    let step;
    try {
      step = await iterator.next();
    } catch (err) {
      throw sourceFailure(err, details(), signal);
    }
    if (step === null || typeof step !== 'object') {
      throw sourceFailure(new TypeError('Iterator result is not an object'), details(), signal);
    }
    return step;
  };

  let finished = false;
  try {
    let step = await next();
    writer.start(resolveColumns(sheetSpec.columns, step.done ? [] : [step.value]));   // D3 peek
    while (!step.done) {
      throwIfAborted(signal);
      writer.write(step.value);
      rowsRead++;
      if (rowsRead % batchSize === 0) await waitForCapacity({ wb, sink, signal });   // D7
      step = await next();
    }
    finished = true;
  } finally {
    if (!finished) await closeIterator(iterator);                                     // D9
  }
}
```

Backlog measurement and waiting (D7). **Spike S0 confirms the field names and coverage before anything else is built.**

```js
const len = (n) => (Number.isFinite(n) ? n : 0);
// readable-stream 2.x (archiver-utils' PassThrough around each worksheet) has no length getters (fact 15).
const streamLen = (s) =>
  len(s?.writableLength ?? s?._writableState?.length) + len(s?.readableLength ?? s?._readableState?.length);

/** Bytes produced but not yet accepted by the sink (fact 15). Internal fields; every read defaults to 0. */
function backlogBytes(wb, sink) {
  let total = streamLen(wb?.zip) + streamLen(wb?.zip?._module) + streamLen(sink.writable);
  for (const ws of wb?._worksheets ?? []) {
    const sb = ws?._stream;                                               // not `ws.stream`: that getter opens a stream
    if (!sb) continue;
    for (const buf of sb.buffers ?? []) total += len(buf?.length);       // not yet handed to a pipe
    for (const p of sb.pipes ?? []) total += streamLen(p);               // archiver's PassThrough wrapper (dominant term)
  }
  return total;
}

async function waitForCapacity({ wb, sink, signal }) {
  // Check abort and sink failure on every pass, not only when over the mark, so a dead sink or an
  // abort is noticed within one batch even when the backlog is small.
  const check = () => {
    throwIfAborted(signal);
    if (sink.failure) throw sink.failure;                                // D11: never keep writing to a dead sink
  };
  for (;;) {
    check();
    if (backlogBytes(wb, sink) <= OUTPUT_HIGH_WATER_MARK) break;
    await sleep(5);                                                       // bounded poll; no bare 'drain' wait
  }
  await tick();
  check();
}
```

This cannot deadlock on its own: it never waits for an event that may not come, and each loop re-checks the abort signal and the sink failure. If the sink's consumer simply stops reading (for example a `writeStream()` consumer that never reads), the write waits for as long as the consumer stalls, like any stream pipeline. Abort releases it. Document this.

Sink handling (D10, D11):

```js
function resolveSink(target, { owned = false } = {}) {
  const track = (writable, isOwned) => {
    const sink = { writable, owned: isOwned, failure: null };
    sink.failed = new Promise((_, reject) => {
      writable.once('error', (err) => { sink.failure = err; reject(err); });
    });
    sink.failed.catch(() => {});                        // observed by race/poll; never an unhandled rejection
    return sink;
  };

  if (typeof target === 'string') {
    const fileStream = createWriteStream(target);
    const sink = track(fileStream, true);
    const finished = new Promise((resolve) => fileStream.once('finish', resolve));
    // Plain listener, not events.once(): once() rejects on 'error' and would settle before the handle closes.
    sink.closed = new Promise((resolve) => fileStream.once('close', resolve));
    sink.cleanup = () => finished;
    sink.remove = async () => {
      fileStream.destroy();
      await sink.closed;                                // Windows: rm fails with EPERM while the handle is open
      await rm(target, { force: true });
    };
    return sink;
  }
  if (target && typeof target.write === 'function') {
    // Caller streams are never destroyed by excelio; library-owned ones (the Buffer collector) pass { owned: true }.
    return track(target, owned);
  }
  throw new ExcelIoError(`Invalid write target — expected file path or writable stream`, { code: 'INVALID_SPEC' });
}
```

Ownership is passed explicitly. `writeOneWorkbook` gains a 5th parameter, `sinkOptions = {}`, and calls `resolveSink(target, sinkOptions)`. `writeOneWorkbookToBuffer` calls `writeOneWorkbook(workbookSpec, wbIndex, collector, options, { owned: true })`. `writeSpec` with `to` and `writeSpecStream` pass nothing (caller-owned). `writeSpecStream` keeps its existing `.catch((err) => passthrough.destroy(err))`.

```js
async function abortWorkbook(wb, sink, err) {
  const quiet = (s) => s?.on?.('error', () => {});
  try {
    if (wb) {
      wb.zip?.unpipe?.(sink.writable);                                 // 1. the caller's stream is not ended
      quiet(wb.zip); quiet(wb.zip?._module);
      for (const ws of wb._worksheets ?? []) (ws?._stream?.pipes ?? []).forEach(quiet);   // _stream: the getter would open one
      wb.zip?.abort?.();                                               // 2.
    }
    if (sink.remove) await sink.remove();                              // 4. file: destroy, wait close, rm
    else if (sink.owned) sink.writable.destroy?.();                    // 3. buffer collector
  } catch (cleanupError) {
    try {                                                              // 5. never replace the original error
      if ((typeof err === 'object' && err !== null) || typeof err === 'function') err.cleanupError = cleanupError;
    } catch { /* frozen or exotic error object: keep the original as is */ }
  }
}
```

Helpers:

```js
function resolveSharedStrings(option = 'auto', sheets) {
  if (typeof option === 'boolean') return option;
  return !sheets.some((s) => isIterableSource(s?.data));                  // D6
}

function computeTotalRows(sheets) {
  let total = 0;
  for (const s of sheets) {
    const kind = rowSourceKind(s.data);
    if (kind === 'array') total += s.data.length;
    else if (kind === 'sync' || kind === 'async') {
      if (s.totalRows === undefined) return undefined;                    // D12
      total += s.totalRows;
    }
  }
  return total;
}

const throwIfAborted = (signal) => {
  if (signal?.aborted) throw new ExcelIoError('Aborted', { code: 'ABORTED' });
};
```

### 4.5 `src/excelio.d.ts`

Apply §3 exactly.

### 4.6 Unchanged

`reader.js`, `coerce.js`, `columns.js` (it works because it receives `[firstRow]`), the worker entry files, and the build scripts.

---

## 5. Memory model

- **Input:** one row at a time, plus whatever the source itself buffers (for example a Mongo cursor batch).
- **Shared strings:** off by default for iterable workbooks (D6). This is the dominant term for distinct-string data.
- **Output backlog:** bounded by `OUTPUT_HIGH_WATER_MARK` (8 MiB), plus one batch of rows (`batchSize` × row size) produced between checks, plus exceljs's 64 KB StreamBuf chunks, **if** S0 confirms the measure covers every queue in fact 15.
- **Target:**
  - With `to: filePath` or a writable, bytes flow out as they are compressed.
  - With **no `to`**, the whole file is collected into a Buffer (unchanged). Document that large streamed exports should use `to`.
- **Version coupling:** the backlog measure reads exceljs/archiver internals. They are optional-chained, so a future exceljs degrades to "yield only" rather than failing. Record in the README that the memory bound was measured with exceljs 4.4.0, and keep the dependency range at `^4.4.0`. Re-run §6.2 whenever exceljs is upgraded.

---

## 6. Tests

### 6.1 `src/__tests__/smoke.test.js`

Add these using the existing `await test(name, fn)` style. Round-trip through `excelio.read` unless the case says otherwise.

| # | Case | Assert |
|---|---|---|
| 1 | async generator, 3 rows, explicit columns | round-trip rows equal; header from columns |
| 2 | sync generator (`function*`) | same as 1 |
| 3 | async iterable **without** columns | keys inferred from the first row; first row written (count 3) |
| 4 | `Readable.from([...rows], { objectMode: true })` | rows written |
| 5 | empty async iterable, with columns | header row only; read gives 0 rows |
| 6 | empty async iterable, no columns | sheet exists, no rows, no throw |
| 7 | workbook: array sheet + iterable sheet + array sheet | all three correct and in the original order |
| 8 | `maxRowsPerSheet: 3`, 7 iterable rows | sheets `Data`, `Data (2)`, `Data (3)` with 3/3/1 rows, each with the header |
| 9 | `maxRowsPerSheet: 3`, 7-row **array** | same as 8 |
| 10 | rollover keeps styling | `freezeHeader` and `autoFilter` on continuation sheets (inspect with `new ExcelJS.Workbook().xlsx.load`) |
| 11 | rollover naming with a collision | logical sheets `Data` (5 rows) and `Data (2)` (1 row), `maxRowsPerSheet: 3` → worksheets in order `Data`, `Data (3)`, `Data (2)` (D8 reservation) |
| 12 | array spec names unchanged | logical sheets `A`, `A`, `B` (duplicate names, no rollover) → `A`, `A (2)`, `B`, exactly as 1.0.0 |
| 13 | `beforeWrite` with iterable and rollover | called once per row; `ctx.sheetPart` 1→2; `ctx.rowNumber` restarts at 2; `ctx.sheetIndex` stays the logical index; `ctx.sheetName` is the actual worksheet name |
| 14 | `onProgress` with iterable and a `totalRows` hint, `batchSize: 2` | events carry `sheetName` and `sheetPart`; final event `rowsProcessed === n`, `totalRows` = hint, `sheetName` = last worksheet |
| 15 | `onProgress` with iterable, no hint | `totalRows === undefined` |
| 16 | iterator throws after 2 rows | rejects `SOURCE_ERROR`; `details.rowsRead === 2`; `cause` is the original error |
| 17 | `[Symbol.asyncIterator]()` itself throws | rejects `SOURCE_ERROR` |
| 18 | `next()` resolves to a non-object (custom iterator) | rejects `SOURCE_ERROR` |
| 19 | abort mid-iteration: async generator yields rows with `await sleep(1)` between them, `batchSize: 1`, `AbortController.abort()` from `setTimeout(20)` | rejects `ABORTED`; generator `finally` ran |
| 20 | abort precedence: the source rejects with an `AbortError` *because* the signal aborted | rejects `ABORTED` (not `SOURCE_ERROR`), `cause` = that `AbortError` |
| 21 | `beforeWrite` throws mid-iterable | rejects with that error; generator `finally` ran |
| 22 | failure cleanup, **file target** (iterator throws after rows) | see "Cleanup assertion protocol" below; the file does not exist afterwards; renaming the temp directory succeeds (proves no open handle on Windows) |
| 23 | failure cleanup, **Buffer output** | rejects; the internal collector was destroyed (spy: wrap `PassThrough.prototype.destroy` for the duration of the test, or expose the collector through `__test`); protocol below |
| 24 | failure cleanup, **caller writable** (`PassThrough` consumed by the test) | rejects; the PassThrough is **not** destroyed (`destroyed === false`) and has **not** emitted `finish`; protocol below |
| 25 | failure cleanup, **`writeStream()`** | the returned stream emits `error` with `SOURCE_ERROR`; protocol below |
| 26 | early sink failure: `to: '<nonexistent dir>/x.xlsx'` with an **endless** async generator (yields a row every 1 ms, never finishes), `batchSize: 10` | rejects within 2 s with the `ENOENT` error, which proves the failure interrupts production; no hang; generator `finally` ran |
| 27 | `useWorker: true` + iterable | succeeds and rows round-trip, which proves inline execution: a worker could not receive a generator, so the call would fail with `WORKER_ERROR` (same outcome-based style as existing worker tests). `logger.warn` called exactly once |
| 28 | `useWorker: 'auto'` + 60,000-row iterable | succeeds inline |
| 29 | invalid `data`: `'abc'`, `{}`, `42`, `true` | `INVALID_SPEC` naming the sheet |
| 30 | the same generator object used by two sheets | `INVALID_SPEC` naming both sheets |
| 31 | the same **array** used by two sheets | allowed; both sheets full |
| 32 | `maxRowsPerSheet: 0`, `1.5`, `2_000_000` | `INVALID_SPEC` |
| 33 | `sharedStrings: 'yes'` | `INVALID_SPEC` |
| 34 | `batchSize: 0`, `-1`, `1.5`, `Infinity` | `INVALID_SPEC` |
| 35 | invalid spec with `useWorker: true` | rejects `INVALID_SPEC` **before** routing: the same error code as inline, no `WORKER_ERROR` (D4) |
| 36 | `sharedStrings` modes, inspected with `jszip` (devDependency, §6.4) | array-only workbook (`'auto'`): `xl/sharedStrings.xml` lists the strings and string cells in `xl/worksheets/sheet1.xml` are `t="s"`; iterable workbook (`'auto'`): string cells are `t="str"`; `true`/`false` override both ways; both read back the same values |
| 37 | `writeStream(spec)` with an iterable | the stream yields a valid workbook |
| 38 | multi-workbook spec, no `to`, an iterable in each | returns `Buffer[]`, both valid |
| 39 | regression | the entire existing suite passes unchanged |

**Cleanup assertion protocol (tests 22–25).**
1. Before the call, register temporary `process.on('unhandledRejection')` and `process.on('uncaughtException')` listeners that record events.
2. After the rejection, wait for any `close` events involved, then `await setTimeout(50)` (several event-loop turns).
3. Assert that nothing was recorded.
4. Remove the listeners in `finally`.
5. The smoke-test process must also **exit by itself** at the end of the run. A leftover handle shows up as a hang, which is the signal that cleanup is incomplete.

### 6.2 New `src/__tests__/memory.test.js` (script `"test:memory": "node --expose-gc src/__tests__/memory.test.js"`)

It is not part of `verify` (Q10), but it is a **release gate** listed in RELEASING.md. In each scenario, sample `process.memoryUsage().rss` every 250 ms and record baseline and peak. In scenarios B and C, also sample `backlogBytes` (exported for tests through `writer.js`'s existing `__test` object) and print its maximum.

| Scenario | Setup | Gate |
|---|---|---|
| A. fast sink | 1,000,000 rows from an `async function*`, 6 columns, every string distinct (`https://example.com/page/${i}?q=${i}`), `to: tmpFile`, default options | peak RSS growth < budget (Q5); the file has 1,000,000 data rows, counted with exceljs's streaming `WorkbookReader` (not `excelio.read`, which loads everything) |
| B. slow sink, single sheet | 300,000 rows into a custom `Writable` (`highWaterMark` 64 KB) whose `write` callback fires after `setTimeout(2)` | peak RSS growth < budget; max `backlogBytes` ≤ high-water mark + one batch |
| C. slow sink, **two large sheets** | two iterable sheets of 200,000 rows each into the slow sink. This catches sheet 2 being produced while sheet 1's zip entry is still compressing (fact 15, the `ws.stream.buffers` term) | same as B |
| D. information only | scenario A with `sharedStrings: true` | print the peak; documents why D6 exists |

Always remove temporary files in `finally`.

### 6.3 Types: `test/types/consumer.mts` and `.cts`

Add:
- a spec whose `data` is `async function*` output, and one whose `data` is an `Iterable<Row>`;
- a sheet with `data` omitted;
- `maxRowsPerSheet` and `sharedStrings: 'auto'`;
- `ctx.sheetPart` read inside `beforeWrite`;
- `new ExcelIoError('x', { code: 'SOURCE_ERROR' })`.

`npm run test:types` must pass.

### 6.4 Dev dependency, package and examples

- Add **`jszip`** as a **devDependency**, matching exceljs's range `^3.10.1`. It is used only by test 36 to read `xl/sharedStrings.xml` and the sheet XML. No runtime dependency is added. Run `npm i -D jszip@^3.10.1` so the lockfile records it.
- Add `examples/10-streaming-rows.mjs`; the file name must match `^\d\d-[a-z-]+\.(mjs|cjs)$`. It writes rows from an `async function*` that simulates a paged API (3 pages × 1,000 rows) with `maxRowsPerSheet: 1200` to the output directory the example runner passes, then prints the worksheet names `Rows`, `Rows (2)`, `Rows (3)`.
- `test/consumer.mjs`: add one async-iterable write against the installed tarball, so the built output is covered.

---

## 7. Documentation

- **README**
  - **Writing:**
    - `data` accepts an array, any iterable, or any async iterable; it may be omitted (empty sheet); iterables are consumed once (one per sheet).
    - Columns are inferred from the first row when omitted, but explicit columns are recommended.
    - Give an example with an async generator, and one line showing a database cursor labelled "any async iterable, e.g. a database cursor; excelio has no database dependency".
  - **New "Large exports" subsection:**
    - Use iterables together with `to` (file or stream); Buffer output still holds the whole file.
    - `sharedStrings` trade-off: on gives a smaller file but memory grows with distinct strings; off gives flat memory and `t="str"` cells.
    - `maxRowsPerSheet` and continuation naming.
    - The `totalRows` hint for progress.
    - The memory bound and the exceljs version it was measured on.
  - **Workers:** iterable data always runs inline; `useWorker: true` is ignored with a warning.
  - **Progress and cancellation:**
    - Iterable sheets check the signal every row and yield every `batchSize` rows, so abort is observed during long iterable writes.
    - A source stuck inside `next()` cannot be interrupted by excelio; pass the same signal to the source.
    - `batchSize` must be a positive integer.
    - The array path is unchanged.
  - **Values and Errors:**
    - New code `SOURCE_ERROR`, plus abort precedence.
    - On failure: a file-path target is removed; a caller's stream is unpiped and neither ended nor destroyed; `err.cleanupError` is set if cleanup itself failed.
    - On success: a caller's stream is ended (unchanged).
  - **Memory and Scope:** rewrite.
    - Arrays still hold all rows.
    - Iterable input plus default `sharedStrings` plus a file or stream target keeps memory bounded; quote the §6.2 numbers.
    - Buffer output collects the file.
    - Reads still load the whole workbook.
- **CHANGELOG** (see Q11): correct the stale `0.1.0 (Unreleased)` heading, then add `## 1.1.0`:
  - **Added:** iterable and async-iterable `data`, `maxRowsPerSheet`, `sharedStrings`, `totalRows` hint, `sheetPart`, `SOURCE_ERROR`.
  - **Fixed:** non-array `data` silently wrote nothing; sheets over 1,048,575 rows produced files Excel cannot open; failed file writes left partial files and open handles; early file errors could hang; `batchSize: 0` silently disabled progress.
  - **Changed:** invalid `batchSize` now throws `INVALID_SPEC`; `SheetSpec.data` is optional in the types.
- **RELEASING.md:** add `npm run test:memory` to the release procedure, and re-run it on every exceljs upgrade.
- **examples/README.md:** list example 10. The file has uncommitted local edits; merge, do not overwrite.

---

## 8. Execution order

0. **Answer §11.** Q1–Q4, Q8, Q9 and Q12 change behaviour or the public API. Get the owner's answers, or proceed with the stated defaults and record that in the PR.
1. **Spike S0: backlog measure.** This is a throwaway script, not committed. Do it before anything else.
   1. Build a workbook with a temporary `async function*` loop, the slow sink of §6.2 B, and two sheets as in C.
   2. Every 100 ms log each term of `backlogBytes` (fact 15) and `rss`, **without** any waiting.
   3. Confirm that:
      - every term exists under the names in fact 15;
      - the sum grows when the sink is slow;
      - the sum accounts for most of the RSS growth.
   4. Then enable `waitForCapacity` and confirm RSS flattens in both the single- and two-sheet cases.
   
   If a term is missing or named differently, fix `backlogBytes` and fact 15 in this plan **before continuing**. If no combination of measurable queues bounds memory, stop and escalate (Q13).
2. `src/rowSource.js`, validation and `preflightWrite` (§4.1–4.3) → tests 29–35.
3. Writer refactor with `createSheetWriter`, name reservation and the array path only → tests 12 and 39. The existing suite must stay green before going further.
4. Rollover → tests 8–11.
5. Iterable path: peek, `waitForCapacity`, source errors, iterator close (§4.4) → tests 1–7, 13–21, 37–38.
6. Sink tracking and the abort sequence (D10, D11) → tests 22–26.
7. `sharedStrings` (D6) plus the `jszip` devDependency → test 36.
8. Worker routing (§4.3) → tests 27–28.
9. Types (§6.3), example 10 and the consumer test (§6.4).
10. Memory gate §6.2 (all scenarios). Set the budget per Q5.
11. Docs (§7), then `npm run verify` on Windows (local). CI covers Linux on Node 22 and 24.

---

## 9. Release

```sh
npm version minor --no-git-tag-version   # 1.0.0 → 1.1.0
npm ci
npm run verify
npm run test:memory
npm run pack:check
npm run publish:dry-run
npm run release                          # explicit owner approval only, per RELEASING.md
```

Then, in `gwp-spidy-v3/node-backend`, run `npm i @dipanshuhandoo/excelio@^1.1.0`. The Spidy Web P4 report uses:
- `data`: one `async function*` per sheet;
- `to`: a temp file;
- `maxRowsPerSheet`: AppConfig `spidy_web_excel_rows_per_sheet`;
- default `sharedStrings`;
- `signal` and `onProgress`.

---

## 10. Acceptance criteria

- Every existing test passes unchanged.
- Array specs with no sheet over 1,048,575 rows produce **semantically equivalent** workbooks to 1.0.0: same sheet names and order, cell values, column widths and formats, styling, and workbook properties. This is compared on parsed content and the relevant unzipped XML, not raw bytes, because zip timestamps differ.
- All §6.1 cases pass on Node 22 and 24, on Linux (CI) and Windows (local run recorded in the PR).
- §6.2 scenarios A–C stay within the agreed budget, and the numbers are recorded in the PR and the README.
- Types compile for ESM and CJS consumers. The installed-tarball consumer test passes.
- A workbook written with `sharedStrings: false` (`t="str"` cells) opens **without a repair prompt** in Microsoft Excel desktop and LibreOffice Calc. This is a manual check recorded in the PR (Q2).
- README, CHANGELOG and RELEASING describe the shipped behaviour. No new **runtime** dependency is added (`jszip` is dev-only).

---

## 11. Open questions (owner decisions)

Each question has a **default** the implementer uses if no answer is given. Record the answers in the PR description, and in the CHANGELOG where relevant.

| # | Question | Why it matters | Default |
|---|---|---|---|
| Q1 | Should `sharedStrings` default to `'auto'` (off when any sheet is iterable), or stay `true` everywhere with callers opting out? | `'auto'` gives bounded memory out of the box. Keeping `true` means every streaming caller must remember `false`, or memory grows with distinct strings (fact 4). Array-only callers see no change either way. | `'auto'` |
| Q2 | Is `t="str"` (fact 17) acceptable for files consumed by Excel and LibreOffice? Is Google Sheets in scope? | `t="str"` is formally the "formula string" type in OOXML. exceljs emits it and it is commonly opened without complaint, but this was **not verified for this package**. If a target rejects it, the options are shared strings (memory cost) or emitting `inlineStr` (requires patching exceljs cell output). | Accept, gated by the manual Excel and LibreOffice check in §10; Google Sheets not in scope |
| Q3 | Should rollover also apply to **array** sheets? | An array over 1,048,575 rows yields an invalid file today; with rollover it yields "Name (2)". It changes behaviour only for inputs that are broken today. The alternative is to reject such arrays with `INVALID_SPEC`. | Roll over |
| Q4 | Should a failed write to a **file path** delete the partial file? | It changes failure behaviour for existing array callers too: today a partial, unreadable file remains, with an open handle. | Delete it |
| Q5 | What peak-RSS growth budget gates the release (§6.2)? | Any number now would be a guess; it depends on exceljs, archiver, zlib and the machine. | Measure in S0 and step 10, set the gate to measured peak + 30 %, and record both numbers |
| Q6 | Should the 8 MiB output high-water mark be a public option? | More API surface; most callers would never change it. | Internal constant in 1.1.0 |
| Q7 | Should `sheetIndex` in `beforeWrite` and progress be the **spec** (logical) index, or the physical worksheet index after rollover? | Callers that switch on `sheetIndex` would see continuation sheets differently. `sheetPart` already distinguishes parts. | Logical index, plus `sheetPart` |
| Q8 | Should reusing the same iterable object in two sheets be rejected (D14), even for reusable collections such as a `Set`? | Rejecting is conservative and prevents the silent empty-sheet trap of single-use iterators, but it also rejects a harmless reused `Set` (the caller must pass a copy). Allowing it requires documenting single-use semantics and accepting silent empty sheets. | Reject, with a message telling the caller to pass a separate iterable or an array |
| Q9 | Release as **1.1.0** (minor)? | It is additive for JavaScript. TypeScript callers with an exhaustive `switch` over `ExcelIoErrorCode` get a compile error from `SOURCE_ERROR`. Invalid `batchSize` values that used to be silently accepted now throw (Q12). | 1.1.0 |
| Q10 | Should `test:memory` also run in CI? | It takes minutes and depends on runner memory, so it could be flaky on shared runners. | Manual release gate only |
| Q11 | `CHANGELOG.md` says `0.1.0 (Unreleased)` while `package.json` and npm say `1.0.0`. What is the correct history? | The 1.1.0 entry needs a correct predecessor. | Rename that entry to `1.0.0` with its publish date |
| Q12 | Should invalid `batchSize` (`0`, negative, fractional, `Infinity`) throw for **all** writes, including the array path? | It now bounds memory and abort latency for iterables. Throwing for arrays too is consistent, but changes behaviour for callers that passed nonsense values (today progress silently stops). The alternative is to validate only when an iterable is present. | Throw for all writes |
| Q13 | If spike S0 shows that no measurable combination of queues bounds memory with a slow sink, what then? | The fallback is to ship iterable input with bounded memory guaranteed only when the sink keeps up (local files, S3 multipart uploads), and document that slow sinks can still buffer. That is weaker than this plan's goal. | Stop and ask the owner; do not ship the weaker guarantee silently |

**Not open** (decided; listed so nobody revisits them): no database adapters in excelio; reads stay non-streaming in 1.1.0; `beforeWrite` stays synchronous; iterable writes never use a worker; continuation sheets are named by the existing duplicate-name rule.
