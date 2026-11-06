import { ExcelIoError } from './errors.js';

/** 'empty' | 'array' | 'sync' | 'async' | 'invalid' */
export function rowSourceKind(data) {
  if (data === undefined || data === null) return 'empty';
  if (Array.isArray(data)) return 'array';
  if (typeof data !== 'object' && typeof data !== 'function') return 'invalid';
  if (typeof data[Symbol.asyncIterator] === 'function') return 'async';
  if (typeof data[Symbol.iterator] === 'function') return 'sync';
  return 'invalid';
}

export const isIterableSource = (data) => {
  const kind = rowSourceKind(data);
  return kind === 'sync' || kind === 'async';
};

export const specHasIterable = (spec) =>
  Array.isArray(spec) && spec.some((wb) => (Array.isArray(wb?.sheets) ? wb.sheets : []).some((s) => isIterableSource(s?.data)));

/** Closes a partially consumed iterator (cursor, stream). Never throws. */
export async function closeIterator(iterator) {
  if (typeof iterator?.return !== 'function') return;
  try { await iterator.return(); } catch { /* the source is already failing or closed */ }
}

/** SOURCE_ERROR, or ABORTED when the caller's signal caused the failure. */
export function sourceFailure(err, details, signal) {
  if (signal?.aborted) return new ExcelIoError('Aborted', { code: 'ABORTED', cause: err, details });
  return new ExcelIoError(`Row source failed for sheet "${details.sheetName}": ${err?.message ?? err}`, {
    code: 'SOURCE_ERROR', cause: err, details
  });
}
