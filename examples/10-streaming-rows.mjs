import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
import path from 'node:path';
import { setImmediate as nextTurn } from 'node:timers/promises';
import { excelio } from '@dipanshuhandoo/excelio';

const output = path.resolve(process.argv[2] ?? 'example-output');
await mkdir(output, { recursive: true });
const file = path.join(output, 'streamed-rows.xlsx');

// Simulates a paginated API; a database cursor or object-mode stream works the same way.
async function fetchPage(page, pageSize) {
  await nextTurn();
  return Array.from({ length: pageSize }, (_, index) => {
    const id = page * pageSize + index + 1;
    return { id, url: `https://example.com/items/${id}`, status: id % 10 === 0 ? 'missing' : 'ok' };
  });
}

async function* pagedRows({ pages, pageSize }) {
  for (let page = 0; page < pages; page++) {
    yield* await fetchPage(page, pageSize);
  }
}

const progress = [];
await excelio.write([{
  workbook: 'streamed',
  sheets: [{
    sheet: 'Rows',
    columns: [{ key: 'id', type: 'number' }, { key: 'url', header: 'URL', width: 40 }, { key: 'status' }],
    styling: { headerBold: true, freezeHeader: true },
    totalRows: 3000,
    data: pagedRows({ pages: 3, pageSize: 1000 })
  }]
}], {
  to: file,
  maxRowsPerSheet: 1200,
  batchSize: 1000,
  onProgress: event => progress.push(`${event.sheetName}: ${event.rowsProcessed}/${event.totalRows}`)
});

const imported = await excelio.read(file);
const names = imported.data[0].sheets.map(sheet => sheet.sheet);
assert.deepEqual(names, ['Rows', 'Rows (2)', 'Rows (3)']);
assert.deepEqual(imported.data[0].sheets.map(sheet => sheet.data.length), [1200, 1200, 600]);
assert.equal(imported.stats.totalRows, 3000);
console.log(`Streamed 3,000 rows from an async generator into ${file}`);
console.log(`Worksheets: ${names.join(', ')}`);
console.log(`Progress: ${progress.join(' | ')}`);
