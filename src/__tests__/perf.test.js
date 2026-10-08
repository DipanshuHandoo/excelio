/**
 * Quick sanity check for file-target writes and a moderately large dataset.
 * Run: npm run test:perf
 */
import assert from 'node:assert/strict';
import { unlinkSync, statSync } from 'node:fs';
import { excelio } from '../index.js';

const tmp = `tmp_excelio_perf_${process.pid}.xlsx`;
const ROWS = 50_000;

console.log(`excelio perf test — ${ROWS} rows\n`);

const data = Array.from({ length: ROWS }, (_, i) => ({
  id: i,
  name: `Customer ${i}`,
  email: `c${i}@example.com`,
  amount: Math.round(Math.random() * 10000) / 100,
  date: new Date(2026, 0, 1 + (i % 365))
}));

const spec = [{
  workbook: 'perf',
  sheets: [{
    sheet: 'Records',
    columns: [
      // header omitted → defaults to key, so round-trip preserves the field names
      { key: 'id',     type: 'number' },
      { key: 'name'   },
      { key: 'email'  },
      { key: 'amount', type: 'number', format: '#,##0.00' },
      { key: 'date',   type: 'date' }
    ],
    styling: { headerBold: true, freezeHeader: true, autoFilter: true },
    data
  }]
}];

const memBefore = process.memoryUsage().rss;

console.log('Writing to file…');
const t0 = Date.now();
await excelio.write(spec, { to: tmp });
const writeMs = Date.now() - t0;
const sizeMB = (statSync(tmp).size / 1024 / 1024).toFixed(2);
console.log(`  ✓ wrote ${sizeMB} MB in ${writeMs}ms`);

console.log('Reading back…');
const t1 = Date.now();
const result = await excelio.read(tmp);
const readMs = Date.now() - t1;
console.log(`  ✓ read ${result.stats.totalRows} rows in ${readMs}ms`);

const memAfter = process.memoryUsage().rss;
const deltaMB = ((memAfter - memBefore) / 1024 / 1024).toFixed(1);
console.log(`  ↑ RSS delta: ${deltaMB} MB`);

assert.equal(result.ok, true);
assert.equal(result.stats.totalRows, ROWS);
assert.equal(result.data[0].sheets[0].data[0].id, 0);
assert.equal(result.data[0].sheets[0].data[ROWS - 1].id, ROWS - 1);

unlinkSync(tmp);
console.log('\n✓ all assertions passed');
