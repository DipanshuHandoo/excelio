import assert from 'node:assert/strict';
import { excelio, ExcelIoError } from '@dipanshuhandoo/excelio';

const spec = [{
  workbook: 'progress',
  sheets: [{
    sheet: 'Invoices',
    columns: [{ key: 'invoice' }, { key: 'amount', type: 'number' }, { key: 'exportedAt', type: 'date' }],
    data: [
      { invoice: 'INV-001', amount: '125.50' },
      { invoice: 'INV-002', amount: '250.00' },
      { invoice: 'INV-003', amount: '75.00' }
    ]
  }]
}];
const progress = [];
const buffer = await excelio.write(spec, {
  useWorker: true,
  batchSize: 1,
  beforeWrite: row => ({ ...row, amount: Number(row.amount), exportedAt: new Date('2026-10-08T00:00:00.000Z') }),
  onProgress(event) {
    progress.push(event);
    console.log(`Write progress: ${event.rowsProcessed}/${event.totalRows}`);
  }
});
assert.equal(progress.at(-1).rowsProcessed, 3);
const result = await excelio.read(buffer, {
  batchSize: 1,
  onProgress: event => console.log(`Read progress: ${event.rowsProcessed} accepted rows`)
});
assert.equal(result.stats.totalRows, 3);
assert.equal(result.data[0].sheets[0].data[0].exportedAt, '2026-10-08T00:00:00.000Z');

const controller = new AbortController();
await assert.rejects(() => excelio.write(spec, {
  signal: controller.signal,
  batchSize: 1,
  onProgress(event) {
    if (event.rowsProcessed === 1) controller.abort();
  }
}), error => error instanceof ExcelIoError && error.code === 'ABORTED');
console.log('Cancellation stopped writing at the next row check; incomplete output was discarded.');
console.log('Hooks and signals force inline execution, even when useWorker is true.');
console.log('Timer-based cancellation cannot interrupt synchronous row loops or initial parsing.');