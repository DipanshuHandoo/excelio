# ExcelIO Examples

These runnable examples use public package exports, synthetic data, and assertions
that fail if expected behavior changes. They live in the repository; they are not
included in the npm tarball. Requires Node.js 22 or later.

## Run from a Checkout

```sh
npm ci
npm run examples
```

The runner builds readable/minified ESM/CommonJS first, runs every example, and
removes temporary XLSX output afterward. It does not publish, use credentials, or
require a web server. Each case runs in its own Node process.

Run a single case by its filename or basename:

```sh
npm run examples -- 03-import-validation
npm run examples -- 07-commonjs-and-minified.cjs
```

## Use Cases

| File | Workflow | Main capabilities |
| --- | --- | --- |
| [01-buffer-roundtrip.mjs](01-buffer-roundtrip.mjs) | Export/import employee data in memory | Buffer output, explicit column keys, coercion, dates, blanks, multiline text |
| [02-sales-report.mjs](02-sales-report.mjs) | Generate a report for people to open in Excel | File target, two sheets, workbook metadata, styling, formats, custom headers, inferred columns |
| [03-import-validation.mjs](03-import-validation.mjs) | Import an uploaded customer workbook | Normalize values, accept partial data, collect row errors, fail-fast validation |
| [04-streaming.mjs](04-streaming.mjs) | Stream an inventory export to a destination | writeStream, pipeline error handling, readable input, inline stream fallback |
| [05-workers-and-concurrency.mjs](05-workers-and-concurrency.mjs) | Export independent tenant workbooks | Forced/automatic workers, configurable thresholds, concurrent jobs, multiple workbook buffers |
| [06-progress-and-cancellation.mjs](06-progress-and-cancellation.mjs) | Observe export/import processing | beforeWrite, explicit added columns, progress, synchronous cancellation checks |
| [07-commonjs-and-minified.cjs](07-commonjs-and-minified.cjs) | Integrate with a CommonJS application | require, minified exports, worker compatibility |
| [08-excel-to-json.mjs](08-excel-to-json.mjs) | Convert an XLSX file to JSON files | All-sheet metadata/data, first-sheet rows, ISO dates, null blanks, JSON serialization |

## Keep Generated Files

Run individual scripts directly after building. File-producing examples default
to the ignored example-output/ directory:

```sh
npm run build
node examples/02-sales-report.mjs
node examples/04-streaming.mjs
node examples/08-excel-to-json.mjs
```

Their optional first argument specifies the output directory. Running a script
again replaces its named demo output file. Do not point it at production data.
Other examples produce console output and do not save files.

The Excel-to-JSON example creates a synthetic customers.xlsx file, then writes
customers-workbook.json with all sheets and customers-rows.json with only the first
sheet's row objects. It reads both JSON files back to verify their contents. To
convert an existing file in your own application, call `excelio.read(inputPath)`
and serialize `result.data` or `result.data[0].sheets[0].data`; the demo generates
its own input so it is safe to run without private data. JSON.stringify omits the
undefined workbook label; sheet names and column metadata are retained in the
all-sheet output.

For use in another project, install `@dipanshuhandoo/excelio` and place the desired
example script inside that project. Its import/require uses the installed package;
the first public release must exist before registry installation is available.

## Important Boundaries

- Keep header equal to key for machine round-trips. Human-friendly headers become
  the property names when files are read back.
- writeStream streams only the first workbook and always runs inline. A writable
  file destination in the pipeline can be replaced with an HTTP response or other
  compatible destination; use the same error/lifecycle handling.
- Stream input does not make reading constant-memory: the whole workbook is loaded.
  Input arrays, shared strings, and buffer output also consume memory.
- Each worker job creates a new thread. Promise.all is application-level concurrency,
  not a reusable pool. Limit concurrent exports according to available resources.
- Worker payloads must be structured-cloneable. Hooks, signals, loggers, and streams
  force inline execution. No callbacks are passed in the worker-only example.
- The cancellation example aborts synchronously inside a progress callback. A timer
  cannot fire during a synchronous row loop; cancellation does not interrupt initial
  parsing or workbook commit. Discard incomplete output after errors or cancellation.
- Minification does not shrink ExcelJS dependencies or guarantee faster processing.