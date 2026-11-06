import { PassThrough } from 'node:stream';
import { createWriteStream } from 'node:fs';
import { rm } from 'node:fs/promises';
import { setImmediate as tick, setTimeout as sleep } from 'node:timers/promises';
import ExcelJS from 'exceljs';

import { ExcelIoError } from './errors.js';
import { resolveColumns } from './columns.js';
import { valueForWrite } from './coerce.js';
import { rowSourceKind, isIterableSource, closeIterator, sourceFailure } from './rowSource.js';

const DEFAULT_BATCH = 5000;
const SHEET_NAME_MAX = 31;
export const EXCEL_MAX_DATA_ROWS = 1_048_575;
// Uncompressed bytes allowed to queue inside exceljs/archiver before an iterable sheet waits.
const OUTPUT_HIGH_WATER_MARK = 8 * 1024 * 1024;
const BACKLOG_POLL_MS = 5;

/**
 * Write one or more workbooks from a spec.
 *
 * Returns:
 *   - When `to` is omitted: a Buffer for a single-workbook spec, or an array of
 *     Buffers (one per workbook) for a multi-workbook spec.
 *   - When `to` is a file path: writes the file, returns the path. Supported only
 *     when spec has exactly one workbook.
 *   - When `to` is a Writable stream: pipes the workbook into it. Single workbook only.
 *
 * Multi-workbook with a single `to` target isn't supported (no zip semantics here).
 * For multi-workbook output, omit `to` and get buffers back.
 */
export async function writeSpec(spec, options = {}) {
  preflightWrite(spec, options);

  const { to } = options;
  if (to !== undefined && spec.length > 1) {
    throw new ExcelIoError(
      `Multi-workbook output requires omitting "to" — caller writes each buffer itself.`,
      { code: 'INVALID_SPEC' }
    );
  }

  if (to !== undefined) {
    await writeOneWorkbook(spec[0], 0, to, options);
    return to;
  }

  const buffers = [];
  for (let wbIndex = 0; wbIndex < spec.length; wbIndex++) {
    const buf = await writeOneWorkbookToBuffer(spec[wbIndex], wbIndex, options);
    buffers.push(buf);
  }
  // Ergonomic: single-workbook spec returns a single Buffer; multi returns Buffer[]
  return buffers.length === 1 ? buffers[0] : buffers;
}

/**
 * Returns a Readable stream containing the first workbook's bytes. If the spec has
 * multiple workbooks, only the first is streamed (callers wanting multi should
 * call writeSpec without `to` to get buffers).
 */
export function writeSpecStream(spec, options = {}) {
  preflightWrite(spec, options);
  const passthrough = new PassThrough();
  writeOneWorkbook(spec[0], 0, passthrough, options).catch(err => passthrough.destroy(err));
  return passthrough;
}

/** Validates the spec and write options. Runs before worker routing and again inside writeSpec. */
export function preflightWrite(spec, options = {}) {
  validateSpec(spec);
  validateWriteOptions(options);
}

/* ────── internals ─────────────────────────────────────────────────────────── */

async function writeOneWorkbookToBuffer(workbookSpec, wbIndex, options) {
  const chunks = [];
  const collector = new PassThrough();
  collector.on('data', (chunk) => chunks.push(chunk));
  await writeOneWorkbook(workbookSpec, wbIndex, collector, options, { owned: true });
  return Buffer.concat(chunks);
}

async function writeOneWorkbook(workbookSpec, wbIndex, target, options, sinkOptions = {}) {
  const { signal, onProgress, beforeWrite, batchSize = DEFAULT_BATCH, logger } = options;
  const maxRowsPerSheet = options.maxRowsPerSheet ?? EXCEL_MAX_DATA_ROWS;
  const sheets = workbookSpec.sheets || [];
  const sink = resolveSink(target, sinkOptions);
  let wb = null;

  try {
    wb = new ExcelJS.stream.xlsx.WorkbookWriter({
      stream: sink.writable,
      useStyles: true,
      useSharedStrings: resolveSharedStrings(options.sharedStrings, sheets)
    });
    applyWorkbookProperties(wb, workbookSpec.properties);

    // Reserve every logical sheet's first-part name up front so array specs keep 1.0.0 names
    // and continuation sheets never take a name a later logical sheet expects.
    const used = new Set();
    const firstNames = sheets.map((s, i) => sanitiseSheetName(s.sheet || `Sheet ${i + 1}`, used, logger));
    const totalRows = computeTotalRows(sheets);
    const state = { rowsProcessed: 0, last: { sheetIndex: 0, sheetName: firstNames[0], sheetPart: 1 } };
    const emit = () => onProgress?.({
      phase: 'write', workbookIndex: wbIndex, ...state.last, rowsProcessed: state.rowsProcessed, totalRows
    });

    for (let sIndex = 0; sIndex < sheets.length; sIndex++) {
      throwIfAborted(signal);
      const sheetSpec = sheets[sIndex];
      const writer = createSheetWriter({
        wb, workbookSpec, wbIndex, sIndex, sheetSpec, firstName: firstNames[sIndex], used, beforeWrite, maxRowsPerSheet,
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

    const committing = wb.commit();
    committing.catch(() => {});                       // a late rejection after the sink failed is not unhandled
    await Promise.race([committing, sink.failed]);
    if (sink.cleanup) await Promise.race([sink.cleanup(), sink.failed]);
    sink.release();
    emit();
  } catch (err) {
    await abortWorkbook(wb, sink, err);
    throw err;
  }
}

/** Writes rows for one logical sheet, rolling over to "<name> (n)" worksheets at maxRowsPerSheet. */
function createSheetWriter({ wb, workbookSpec, wbIndex, sIndex, sheetSpec, firstName, used, beforeWrite, maxRowsPerSheet, onOpen, onRow }) {
  const baseName = sheetSpec.sheet || `Sheet ${sIndex + 1}`;
  let cols = [];
  let ws = null;
  let sheetName = null;
  let part = 0;
  let rowsInPart = 0;

  const open = () => {
    part += 1;
    rowsInPart = 0;
    // The truncation warning was already logged when the first part's name was reserved.
    sheetName = part === 1 ? firstName : sanitiseSheetName(baseName, used, null);
    // WorksheetWriter (streaming) doesn't allow setting `views` after creation.
    const wsOptions = {};
    if (sheetSpec.styling?.freezeHeader) wsOptions.views = [{ state: 'frozen', ySplit: 1 }];
    ws = wb.addWorksheet(sheetName, wsOptions);
    if (cols.length) {
      ws.columns = cols.map(c => ({
        key: c.key,
        header: c.header,
        width: c.width,
        style: c.format ? { numFmt: c.format } : undefined
      }));
    }
    applyStylingPostColumns(ws, sheetSpec.styling, cols.length);
    onOpen(sheetName, part);
  };

  return {
    get sheetName() { return sheetName ?? firstName; },
    start(resolvedColumns) {
      cols = resolvedColumns;
      open();
    },
    write(row) {
      if (rowsInPart >= maxRowsPerSheet) {
        ws.commit();
        open();
      }
      let out = row;
      if (beforeWrite) {
        out = beforeWrite(row, {
          workbookIndex: wbIndex,
          workbookName: workbookSpec.workbook,
          sheetIndex: sIndex,
          sheetName,
          sheetPart: part,
          rowNumber: rowsInPart + 2  // header is row 1
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

async function writeIterableSheet({ writer, wb, sink, sheetSpec, kind, signal, batchSize, wbIndex, sIndex }) {
  let rowsRead = 0;
  const details = () => ({ workbookIndex: wbIndex, sheetIndex: sIndex, sheetName: writer.sheetName, rowsRead });

  let iterator;
  try {
    iterator = kind === 'async' ? sheetSpec.data[Symbol.asyncIterator]() : sheetSpec.data[Symbol.iterator]();
  } catch (err) {
    throw sourceFailure(err, details(), signal);
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
    // The peeked first row infers columns (when none are declared) and is then written normally.
    writer.start(resolveColumns(sheetSpec.columns, step.done ? [] : [step.value]));
    while (!step.done) {
      throwIfAborted(signal);
      writer.write(step.value);
      rowsRead++;
      if (rowsRead % batchSize === 0) await waitForCapacity({ wb, sink, signal });
      step = await next();
    }
    finished = true;
  } finally {
    if (!finished) await closeIterator(iterator);
  }
}

const len = (n) => (Number.isFinite(n) ? n : 0);
// readable-stream 2.x (archiver-utils' PassThrough around each worksheet) has no length getters.
const streamLen = (s) =>
  len(s?.writableLength ?? s?._writableState?.length) + len(s?.readableLength ?? s?._readableState?.length);

/**
 * Bytes produced but not yet accepted by the sink. Reads exceljs 4.4 / archiver 5.3 internals; every
 * read defaults to 0 so an incompatible version degrades to "yield only" instead of crashing.
 */
function backlogBytes(wb, sink) {
  let total = streamLen(wb?.zip) + streamLen(wb?.zip?._module) + streamLen(sink?.writable);
  for (const ws of wb?._worksheets ?? []) {
    const sb = ws?._stream;                           // not `ws.stream`: that getter opens a stream
    if (!sb) continue;
    for (const buf of sb.buffers ?? []) total += len(buf?.length);
    for (const p of sb.pipes ?? []) total += streamLen(p);
  }
  return total;
}

const backlogStats = { peak: 0 };

async function waitForCapacity({ wb, sink, signal }) {
  const check = () => {
    throwIfAborted(signal);
    if (sink.failure) throw sink.failure;
  };
  for (;;) {
    check();
    const backlog = backlogBytes(wb, sink);
    if (backlog > backlogStats.peak) backlogStats.peak = backlog;
    if (backlog <= OUTPUT_HIGH_WATER_MARK) break;
    await sleep(BACKLOG_POLL_MS);
  }
  await tick();
  check();
}

function resolveSink(target, { owned = false } = {}) {
  const track = (writable, isOwned) => {
    const sink = { writable, owned: isOwned, failure: null };
    let onError;
    sink.failed = new Promise((_, reject) => {
      onError = (err) => {
        if (sink.failure) return;
        sink.failure = err;
        reject(err);
      };
      writable.on('error', onError);
    });
    sink.failed.catch(() => {});
    // Caller-owned streams get their error handling back once excelio is done with them.
    sink.release = () => { if (!isOwned) writable.removeListener('error', onError); };
    return sink;
  };

  if (typeof target === 'string') {
    const fileStream = createWriteStream(target);
    const sink = track(fileStream, true);
    // Register 'finish' and 'close' now so neither is missed before cleanup()/remove() awaits it.
    const finished = new Promise((resolve) => fileStream.once('finish', resolve));
    const closed = new Promise((resolve) => fileStream.once('close', resolve));
    sink.cleanup = () => finished;
    sink.remove = async () => {
      fileStream.destroy();
      await closed;                                   // Windows: rm fails with EPERM while the handle is open
      await rm(target, { force: true });
    };
    return sink;
  }
  if (target && typeof target.write === 'function') {
    return track(target, owned);
  }
  throw new ExcelIoError(`Invalid write target — expected file path or writable stream`, { code: 'INVALID_SPEC' });
}

async function abortWorkbook(wb, sink, err) {
  const quiet = (s) => s?.on?.('error', () => {});
  try {
    if (wb) {
      wb.zip?.unpipe?.(sink.writable);                // a caller's stream must not be ended by the archive
      quiet(wb.zip);
      quiet(wb.zip?._module);
      for (const ws of wb._worksheets ?? []) (ws?._stream?.pipes ?? []).forEach(quiet);
      wb.zip?.abort?.();
    }
    if (sink.remove) await sink.remove();
    else if (sink.owned) sink.writable.destroy?.();
    sink.release();
  } catch (cleanupError) {
    try {
      if ((typeof err === 'object' && err !== null) || typeof err === 'function') err.cleanupError = cleanupError;
    } catch { /* frozen or exotic error object: keep the original as is */ }
  }
}

function resolveSharedStrings(option = 'auto', sheets) {
  if (typeof option === 'boolean') return option;
  return !sheets.some((s) => isIterableSource(s?.data));
}

function computeTotalRows(sheets) {
  let total = 0;
  for (const s of sheets) {
    const kind = rowSourceKind(s.data);
    if (kind === 'array') total += s.data.length;
    else if (kind === 'sync' || kind === 'async') {
      if (s.totalRows === undefined) return undefined;
      total += s.totalRows;
    }
  }
  return total;
}

function throwIfAborted(signal) {
  if (signal?.aborted) throw new ExcelIoError('Aborted', { code: 'ABORTED' });
}

function applyWorkbookProperties(wb, props) {
  if (!props) return;
  if (props.author)   wb.creator = props.author;
  if (props.created)  wb.created = props.created instanceof Date ? props.created : new Date(props.created);
  if (props.modified) wb.modified = props.modified instanceof Date ? props.modified : new Date(props.modified);
  if (props.company)  wb.company = props.company;
  if (props.title)    wb.title = props.title;
}

function applyStylingPostColumns(ws, styling, colCount) {
  if (!styling) return;
  if (styling.headerBold) {
    ws.getRow(1).font = { bold: true };
  }
  if (styling.autoFilter && colCount > 0) {
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: colCount } };
  }
  // freezeHeader is applied at addWorksheet time (see wsOptions in createSheetWriter).
}

const FORBIDDEN_SHEET_CHARS = /[\\/?*[\]:]/g;

function sanitiseSheetName(raw, used, logger) {
  let name = String(raw).replace(FORBIDDEN_SHEET_CHARS, '_');
  if (name.length > SHEET_NAME_MAX) {
    if (logger?.warn) logger.warn(`excelio: sheet name "${raw}" truncated to ${SHEET_NAME_MAX} chars`);
    name = name.slice(0, SHEET_NAME_MAX);
  }
  if (!used.has(name)) { used.add(name); return name; }
  let i = 2;
  while (true) {
    const suffix = ` (${i})`;
    const candidate = (name.length + suffix.length > SHEET_NAME_MAX
      ? name.slice(0, SHEET_NAME_MAX - suffix.length)
      : name) + suffix;
    if (!used.has(candidate)) { used.add(candidate); return candidate; }
    i++;
  }
}

function validateSpec(spec) {
  if (!Array.isArray(spec) || spec.length === 0) {
    throw new ExcelIoError('Spec must be a non-empty array of workbooks', { code: 'INVALID_SPEC' });
  }
  const owners = new Map();                          // iterable object → first sheet name using it
  spec.forEach((wb, wbIndex) => {
    if (!wb || typeof wb !== 'object') {
      throw new ExcelIoError('Workbook entries must be objects', { code: 'INVALID_SPEC' });
    }
    if (!Array.isArray(wb.sheets) || wb.sheets.length === 0) {
      throw new ExcelIoError('Workbook must have at least one sheet', { code: 'INVALID_SPEC' });
    }
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
    throw new ExcelIoError('batchSize must be a positive integer', { code: 'INVALID_SPEC' });
  }
}

// Helpers exported for tests
export const __test = { sanitiseSheetName, backlogBytes, backlogStats, OUTPUT_HIGH_WATER_MARK };
