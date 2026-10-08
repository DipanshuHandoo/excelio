import { Readable, PassThrough } from 'node:stream';
import { createWriteStream } from 'node:fs';
import { pipeline } from 'node:stream/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';

import { ExcelIoError } from './errors.js';
import { resolveColumns } from './columns.js';
import { valueForWrite } from './coerce.js';

const DEFAULT_BATCH = 5000;
const SHEET_NAME_MAX = 31;

/**
 * Write one or more workbooks from a spec.
 *
 * Returns:
 *   - When `to` is omitted: an array of Buffers, one per workbook (single-element
 *     array when spec has one workbook).
 *   - When `to` is a file path: writes the file, returns the path. Supported only
 *     when spec has exactly one workbook.
 *   - When `to` is a Writable stream: pipes the workbook into it. Single workbook only.
 *
 * Multi-workbook with a single `to` target isn't supported (no zip semantics here).
 * For multi-workbook output, omit `to` and get buffers back.
 */
export async function writeSpec(spec, options = {}) {
  validateSpec(spec);

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
  validateSpec(spec);
  const passthrough = new PassThrough();
  writeOneWorkbook(spec[0], 0, passthrough, options).catch(err => passthrough.destroy(err));
  return passthrough;
}

/* ────── internals ─────────────────────────────────────────────────────────── */

async function writeOneWorkbookToBuffer(workbookSpec, wbIndex, options) {
  const chunks = [];
  const collector = new PassThrough();
  collector.on('data', (chunk) => chunks.push(chunk));
  await writeOneWorkbook(workbookSpec, wbIndex, collector, options);
  return Buffer.concat(chunks);
}

async function writeOneWorkbook(workbookSpec, wbIndex, target, options) {
  const { signal, onProgress, beforeWrite, batchSize = DEFAULT_BATCH, logger } = options;

  const sink = resolveSink(target);

  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
    stream: sink.writable,
    useStyles: true,
    useSharedStrings: true
  });

  applyWorkbookProperties(wb, workbookSpec.properties);

  const usedNames = new Set();
  const sheets = workbookSpec.sheets || [];
  let totalRows = sheets.reduce((n, s) => n + (s.data?.length || 0), 0);
  let rowsProcessed = 0;

  for (let sIndex = 0; sIndex < sheets.length; sIndex++) {
    if (signal?.aborted) throw new ExcelIoError('Aborted', { code: 'ABORTED' });

    const sheetSpec = sheets[sIndex];
    const sheetName = sanitiseSheetName(sheetSpec.sheet || `Sheet ${sIndex + 1}`, usedNames, logger);

    // Pass styling-related options at creation time — WorksheetWriter (streaming)
    // doesn't allow setting `views` after the fact.
    const wsOptions = {};
    if (sheetSpec.styling?.freezeHeader) {
      wsOptions.views = [{ state: 'frozen', ySplit: 1 }];
    }
    const ws = wb.addWorksheet(sheetName, wsOptions);

    const cols = resolveColumns(sheetSpec.columns, sheetSpec.data);
    if (cols.length) {
      ws.columns = cols.map(c => ({
        key: c.key,
        header: c.header,
        width: c.width,
        style: c.format ? { numFmt: c.format } : undefined
      }));
    }

    applyStylingPostColumns(ws, sheetSpec.styling, cols.length);

    const data = sheetSpec.data || [];
    for (let i = 0; i < data.length; i++) {
      if (signal?.aborted) throw new ExcelIoError('Aborted', { code: 'ABORTED' });

      let row = data[i];
      if (beforeWrite) {
        row = beforeWrite(row, {
          workbookIndex: wbIndex,
          workbookName: workbookSpec.workbook,
          sheetIndex: sIndex,
          sheetName,
          rowNumber: i + 2  // header is row 1
        }) ?? row;
      }

      const rowValues = cols.length
        ? cols.reduce((acc, c) => { acc[c.key] = valueForWrite(row[c.key], c.type); return acc; }, {})
        : row;
      ws.addRow(rowValues).commit();

      rowsProcessed++;
      if (onProgress && rowsProcessed % batchSize === 0) {
        onProgress({ phase: 'write', workbookIndex: wbIndex, sheetIndex: sIndex, rowsProcessed, totalRows });
      }
    }

    ws.commit();
  }

  await wb.commit();
  if (sink.cleanup) await sink.cleanup();

  if (onProgress) {
    onProgress({ phase: 'write', workbookIndex: wbIndex, sheetIndex: sheets.length - 1, rowsProcessed, totalRows });
  }
}

function resolveSink(target) {
  // File path — register the close promise immediately so we don't miss the
  // 'finish' event if it fires before cleanup() is awaited.
  if (typeof target === 'string') {
    const fileStream = createWriteStream(target);
    const closed = new Promise((resolve, reject) => {
      fileStream.once('finish', resolve);
      fileStream.once('error', reject);
    });
    return { writable: fileStream, cleanup: () => closed };
  }
  // Writable stream supplied by caller — they own its lifecycle.
  if (target && typeof target.write === 'function') {
    return { writable: target, cleanup: null };
  }
  throw new ExcelIoError(`Invalid write target — expected file path or writable stream`, { code: 'INVALID_SPEC' });
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
  // freezeHeader is applied at addWorksheet time (see wsOptions in writeOneWorkbook).
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
  for (const wb of spec) {
    if (!wb || typeof wb !== 'object') {
      throw new ExcelIoError('Workbook entries must be objects', { code: 'INVALID_SPEC' });
    }
    if (!Array.isArray(wb.sheets) || wb.sheets.length === 0) {
      throw new ExcelIoError('Workbook must have at least one sheet', { code: 'INVALID_SPEC' });
    }
  }
}

// Helpers exported for tests
export const __test = { sanitiseSheetName };
