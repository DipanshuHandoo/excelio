/**
 * Cell value coercion in both directions.
 *
 * Write side: accept a JS value + optional column type, normalise to what exceljs
 * expects. Excel-native types are Number, String, Boolean, Date, formula objects,
 * and null.
 *
 * Read side: take a cell from exceljs and turn it into a portable JS value. Dates
 * default to ISO 8601 strings; configurable via `dateFormat`.
 */

const ISO_DATE_PATTERN = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?)?$/;

/**
 * Normalise a JS value for writing into a cell. Returns the value as-is for most
 * types; converts ISO date strings + epoch ms to Date when the column declares
 * `type: 'date'`.
 */
export function valueForWrite(value, columnType) {
  if (value === null || value === undefined) return null;

  if (columnType === 'date') {
    if (value instanceof Date) return value;
    if (typeof value === 'number') return new Date(value);
    if (typeof value === 'string' && ISO_DATE_PATTERN.test(value)) {
      const d = new Date(value);
      return isNaN(d.getTime()) ? value : d;
    }
    return value;
  }

  if (columnType === 'number' && typeof value === 'string' && value !== '') {
    const n = Number(value);
    return Number.isFinite(n) ? n : value;
  }

  if (columnType === 'boolean') {
    if (typeof value === 'boolean') return value;
    if (value === 'true' || value === 1) return true;
    if (value === 'false' || value === 0) return false;
  }

  return value;
}

/**
 * Convert an exceljs cell value into a portable JS value.
 *
 * options:
 *   nullForBlank: boolean   — empty cells become null (default true)
 *   trimStrings:  boolean   — strip leading/trailing whitespace from strings
 *   dateFormat:   'iso' | 'date' | 'epoch'
 */
export function valueForRead(cell, { nullForBlank = true, trimStrings = true, dateFormat = 'iso' } = {}) {
  if (cell === undefined || cell === null) return nullForBlank ? null : '';

  // exceljs returns CellValue which can be primitive or object; cell.value is the
  // canonical accessor. When called with a raw value (e.g. from streaming), it's
  // already the primitive.
  const raw = (cell && typeof cell === 'object' && 'value' in cell) ? cell.value : cell;

  if (raw === null || raw === undefined || raw === '') return nullForBlank ? null : '';

  // Date cells
  if (raw instanceof Date) return formatDate(raw, dateFormat);

  // Hyperlink object: { text, hyperlink, ... } — return the visible text
  if (typeof raw === 'object' && raw.text !== undefined && raw.hyperlink !== undefined) {
    return trimStrings ? String(raw.text).trim() : String(raw.text);
  }

  // Rich text object: { richText: [{ text, font }, ...] }
  if (typeof raw === 'object' && Array.isArray(raw.richText)) {
    const joined = raw.richText.map(r => r.text).join('');
    return trimStrings ? joined.trim() : joined;
  }

  // Formula object: { formula, result }
  if (typeof raw === 'object' && raw.formula !== undefined) {
    return raw.result !== undefined ? raw.result : null;
  }

  // Error cell: { error: '#N/A' }
  if (typeof raw === 'object' && raw.error !== undefined) return null;

  if (typeof raw === 'string') return trimStrings ? raw.trim() : raw;
  return raw;
}

function formatDate(d, fmt) {
  switch (fmt) {
    case 'date':  return d;
    case 'epoch': return d.getTime();
    case 'iso':
    default:      return d.toISOString();
  }
}
