import assert from 'node:assert/strict';
import { createReadStream, createWriteStream } from 'node:fs';
import { mkdir } from 'node:fs/promises';
import path from 'node:path';
import { pipeline } from 'node:stream/promises';
import { excelio } from '@dipanshuhandoo/excelio';

const output = path.resolve(process.argv[2] ?? 'example-output');
await mkdir(output, { recursive: true });
const file = path.join(output, 'streamed-inventory.xlsx');
const spec = [{
  workbook: 'inventory',
  sheets: [{
    sheet: 'Stock',
    data: Array.from({ length: 200 }, (_, index) => ({
      sku: `SKU-${String(index + 1).padStart(4, '0')}`,
      quantity: index % 25,
      warehouse: index % 2 === 0 ? 'North' : 'South'
    }))
  }]
}];

await pipeline(excelio.writeStream(spec), createWriteStream(file));
const imported = await excelio.read(createReadStream(file), { useWorker: true });
assert.equal(imported.ok, true);
assert.equal(imported.stats.totalRows, 200);
assert.equal(imported.data[0].sheets[0].data[0].sku, 'SKU-0001');
console.log(`Streamed export saved: ${file}`);
console.log(`Read ${imported.stats.totalRows} rows from a readable stream.`);
console.log('Stream input stays inline; reading still loads the complete workbook.');