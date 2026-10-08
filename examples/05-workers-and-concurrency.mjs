import assert from 'node:assert/strict';
import { Buffer } from 'node:buffer';
import { excelio } from '@dipanshuhandoo/excelio';

function tenantWorkbook(tenant) {
  return {
    workbook: tenant,
    sheets: [{
      sheet: 'Orders',
      data: Array.from({ length: 500 }, (_, index) => ({
        tenant, orderId: `${tenant}-${index + 1}`, amount: (index + 1) * 10
      }))
    }]
  };
}

const north = tenantWorkbook('north');
const south = tenantWorkbook('south');
const forced = await excelio.write([north], { useWorker: true });
assert.ok(Buffer.isBuffer(forced));
assert.equal((await excelio.read(forced, { useWorker: true })).stats.totalRows, 500);

const automatic = await excelio.write([north], { workerThreshold: { rows: 500 } });
assert.ok(Buffer.isBuffer(automatic));
const imported = await excelio.read(automatic, { workerThreshold: { bytes: automatic.length } });
assert.equal(imported.stats.totalRows, 500);

const separateExports = await Promise.all([
  excelio.write([north], { useWorker: true }),
  excelio.write([south], { useWorker: true })
]);
assert.equal(separateExports.length, 2);
assert.ok(separateExports.every(buffer => Buffer.isBuffer(buffer)));

const multipleWorkbooks = await excelio.write([north, south], { useWorker: true });
assert.equal(multipleWorkbooks.length, 2);
assert.ok(multipleWorkbooks.every(buffer => Buffer.isBuffer(buffer)));
console.log('Forced workers, automatic thresholds, two concurrent jobs, and Buffer[] output passed.');
console.log('Promise.all submits independent jobs; each gets a fresh worker, not a pooled or split workbook.');