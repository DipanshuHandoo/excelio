import { Readable } from 'node:stream';
import { Buffer } from 'node:buffer';
import { createReadStream } from 'node:fs';
import { stat } from 'node:fs/promises';
import ExcelJS from 'exceljs';

import { ExcelIoError } from './errors.js';
import { valueForRead } from './coerce.js';

const DEFAULT_BATCH = 5000;

/**
 * Read a workbook from a Buffer, file path, or readable stream.
 *
 * Currently uses the non-streaming `xlsx.load` / `xlsx.read` API for reliability.
 * Large files may exceed available memory. Worker-thread offload keeps parsing
 * off the main thread, but does not reduce total process memory or isolate it.
 * A future optimisation can swap to `stream.xlsx.WorkbookReader` once exceljs
 * stabilises its async iterator semantics across versions.
 */
export async function readSource(input, options = {}) {
  const startedAt = Date.now();
  const stats = { totalRows: 0, totalSheets: 0, durationMs: 0 };
  const errors = [];

  const {
    headerRow = 1,
    nullForBlank = true,
    trimStrings = true,
    dateFormat = 'iso',
    validateRow,
    afterRead,
    onRowError = 'collect',
    batchSize = DEFAULT_BATCH,
    onProgress,
    signal
  } = options;

  const wb = new ExcelJS.Workbook();
  try {
    if (Buffer.isBuffer(input)) {
      await wb.xlsx.load(input);
    } else if (input instanceof Uint8Array) {
      // Worker boundary downgrades Buffer to Uint8Array; rewrap.
      await wb.xlsx.load(Buffer.from(input.buffer, input.byteOffset, input.byteLength));
    } else if (typeof input === 'string') {
      await wb.xlsx.readFile(input);
    } else if (input && typeof input.pipe === 'function') {
      await wb.xlsx.read(input);
    } else {
      throw new ExcelIoError('Invalid input — expected Buffer, file path, or readable stream', { code: 'INVALID_SPEC' });
    }
  } catch (err) {
    if (err instanceof ExcelIoError) throw err;
    throw new ExcelIoError(`Failed to parse workbook: ${err.message}`, { code: 'IO_ERROR', cause: err });
  }

  const sheets = [];

  for (let s = 0; s < wb.worksheets.length; s++) {
    if (signal?.aborted) throw new ExcelIoError('Aborted', { code: 'ABORTED' });

    const ws = wb.worksheets[s];
    stats.totalSheets++;

    const sheetMeta = { sheet: ws.name, columns: [], data: [] };
    let columnKeys = null;

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      // Header row
      if (rowNumber === headerRow) {
        columnKeys = readHeaderRow(row);
        sheetMeta.columns = columnKeys
          .filter(Boolean)
          .map(k => ({ key: k, header: k }));
        return;
      }
      if (rowNumber < headerRow || !columnKeys) return;
      if (signal?.aborted) throw new ExcelIoError('Aborted', { code: 'ABORTED' });

      const obj = {};
      let isEmpty = true;
      for (let i = 0; i < columnKeys.length; i++) {
        const key = columnKeys[i];
        if (!key) continue;
        const cell = row.getCell(i + 1);
        const v = valueForRead(cell, { nullForBlank, trimStrings, dateFormat });
        if (v !== null && v !== '' && v !== undefined) isEmpty = false;
        obj[key] = v;
      }
      if (isEmpty) return;

      const ctx = {
        workbookIndex: 0,
        workbookName: undefined,
        sheetIndex: s,
        sheetName: ws.name,
        rowNumber
      };

      let outRow = obj;
      if (afterRead) {
        try {
          outRow = afterRead(obj, ctx) ?? obj;
        } catch (err) {
          const e = mkRowError(ctx, 'TRANSFORM_ERROR', err.message);
          if (onRowError === 'fail') throw new ExcelIoError(e.message, { code: 'ROW_VALIDATION', details: e });
          errors.push(e);
          return;
        }
      }

      if (validateRow) {
        const issue = validateRow(outRow, ctx);
        if (issue) {
          const e = mkRowError(ctx, issue.code || 'VALIDATION_FAILED', issue.message, issue.column);
          if (onRowError === 'fail') throw new ExcelIoError(e.message, { code: 'ROW_VALIDATION', details: e });
          errors.push(e);
          return;
        }
      }

      sheetMeta.data.push(outRow);
      stats.totalRows++;

      if (onProgress && stats.totalRows % batchSize === 0) {
        onProgress({ phase: 'read', sheetName: ws.name, rowsProcessed: stats.totalRows });
      }
    });

    sheets.push(sheetMeta);
  }

  stats.durationMs = Date.now() - startedAt;

  return {
    ok: errors.length === 0,
    data: [{ workbook: undefined, sheets }],
    errors,
    stats
  };
}

/* ────── helpers ───────────────────────────────────────────────────────────── */

function readHeaderRow(row) {
  const headers = [];
  // exceljs row.values is sparse: index 0 unused, headers start at 1
  const values = row.values;
  if (!Array.isArray(values)) return headers;
  for (let i = 1; i < values.length; i++) {
    const v = values[i];
    headers.push(v == null ? null : String(v).trim());
  }
  return headers;
}

function mkRowError(ctx, code, message, column) {
  return {
    workbook: ctx.workbookIndex,
    sheet: ctx.sheetName,
    row: ctx.rowNumber,
    column: column ?? null,
    code,
    message
  };
}
