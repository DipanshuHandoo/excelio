/**
 * excelio — Plug-and-Play Excel ↔ JSON utility.
 *
 * Symmetric API. Works on Buffers, file paths, or streams. Streams everything
 * internally so 1M-row datasets stay memory-safe (~50 MB constant). Worker-thread
 * offload above a row/byte threshold (auto by default).
 *
 * The whole package (this directory) is self-contained — only depends on `exceljs`
 * and Node built-ins. Drop the directory into any Node project and it works.
 *
 * Quick start:
 *
 *   import { excelio } from '@/shared/utils/excelio';
 *
 *   // Write
 *   const buf = await excelio.write([{
 *     workbook: 'sales',
 *     sheets: [{ sheet: 'Q1', data: [{ name: 'Alice', amount: 100 }] }]
 *   }]);
 *
 *   // Read
 *   const result = await excelio.read(buf, {
 *     validateRow: (row) => row.amount < 0 ? { message: 'negative amount' } : null
 *   });
 *
 * See README.md for the full reference.
 */

import { writeSpec, writeSpecStream } from './writer.js';
import { readSource } from './reader.js';
import { ExcelIoError } from './errors.js';
import { maybeUseWorker } from './workerPool.js';

export const excelio = {
  /**
   * Write one or more workbooks. Returns a Buffer (or array of Buffers for
   * multi-workbook specs) when `to` is omitted; writes to file/stream when `to`
   * is provided and returns the target.
   */
  async write(spec, options = {}) {
    return maybeUseWorker('write', spec, options, () => writeSpec(spec, options));
  },

  /**
   * Returns a Readable stream for the first workbook. Synchronous return;
   * errors during the underlying write surface via the stream's 'error' event.
   */
  writeStream(spec, options = {}) {
    return writeSpecStream(spec, options);
  },

  /**
   * Read a workbook from a Buffer, file path, or Readable stream. Returns a
   * structured result with parsed data, per-row errors, and stats.
   */
  async read(input, options = {}) {
    return maybeUseWorker('read', input, options, () => readSource(input, options));
  }
};

export { ExcelIoError };
export default excelio;
