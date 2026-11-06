# Changelog

## 1.1.0 (2026-11-06)

### Added

- Sheet `data` accepts any iterable or async iterable (generators, database cursors,
  object-mode streams) as well as arrays; iterable sheets are written inline with
  output backlog bounded at about 8 MiB, so memory stays flat for large exports.
- `maxRowsPerSheet` write option: sheets continue on `"<name> (2)"`, `"<name> (3)"`, ….
- `sharedStrings: 'auto' | true | false` write option (`'auto'`: off when a workbook
  has an iterable sheet).
- `totalRows` sheet hint for progress, and `sheetPart` in `beforeWrite` context and
  write progress events (which now also carry `sheetName`).
- `SOURCE_ERROR` error code for failing row sources.
- `npm run test:memory` release gate.

### Fixed

- Non-array `data` (for example a generator) silently wrote an empty sheet.
- Sheets over 1,048,575 rows produced files Excel cannot open; they now roll over.
- Failed writes to a file path left a partial file and an open handle; the file is
  now removed. Failed writes to a caller's stream no longer end that stream.
- Early file errors (such as a missing directory) could hang the write.
- `batchSize: 0` silently disabled progress events.

### Changed

- Invalid `batchSize` (not a positive integer) now throws `INVALID_SPEC`.
- Specs and write options are validated before worker routing.
- `SheetSpec.data` is optional in the type declarations. TypeScript code with an
  exhaustive switch over `ExcelIoErrorCode` must handle `SOURCE_ERROR`.

## 1.0.0 (2026-10-08)

- Package as `@dipanshuhandoo/excelio`, supporting Node.js 22 and 24 LTS.
- Add readable/minified ESM and CommonJS builds, worker bundles, and declarations.
- Preserve Buffer results across worker threads and keep stream inputs inline.
- Add installed-tarball consumer verification and explicit public npm commands.
- Add runnable use-case examples, including XLSX-to-JSON conversion, selection/cleanup
	tooling, and CI coverage.
- Document XLSX API behavior, memory limitations, and release/support policies.