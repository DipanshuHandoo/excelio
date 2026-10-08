/**
 * Column metadata helpers.
 *
 * A column entry shape:
 *   { key, header?, type?, format?, width? }
 *
 *   key    : property name on each data row (required)
 *   header : displayed in the first row (defaults to key)
 *   type   : 'string' | 'number' | 'boolean' | 'date' (informs cell formatting)
 *   format : Excel number-format string (e.g. 'yyyy-mm-dd', '#,##0.00')
 *   width  : column width in Excel units (default 15)
 */

/**
 * Resolve the columns for a sheet. If the caller provided columns, validate/normalise.
 * Otherwise infer from the first row's keys (preserving insertion order).
 */
export function resolveColumns(declaredColumns, data) {
  if (Array.isArray(declaredColumns) && declaredColumns.length) {
    return declaredColumns.map(normaliseColumn);
  }
  if (!data?.length) return [];
  const firstRow = data[0] || {};
  return Object.keys(firstRow).map(k => normaliseColumn({ key: k }));
}

function normaliseColumn(c) {
  if (!c || typeof c !== 'object' || !c.key) {
    throw new Error(`excelio: column entry must have a "key" property; got ${JSON.stringify(c)}`);
  }
  return {
    key:    String(c.key),
    header: c.header ?? c.key,
    type:   c.type ?? null,
    format: c.format ?? defaultFormatFor(c.type),
    width:  c.width ?? 15
  };
}

function defaultFormatFor(type) {
  switch (type) {
    case 'date':   return 'yyyy-mm-dd hh:mm:ss';
    case 'number': return null;
    default:       return null;
  }
}
