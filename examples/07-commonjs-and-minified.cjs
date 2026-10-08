const assert = require('node:assert/strict');
const { Buffer } = require('node:buffer');
const { excelio } = require('@dipanshuhandoo/excelio');
const { excelio: minified } = require('@dipanshuhandoo/excelio/min');

async function main() {
  const spec = [{
    workbook: 'products',
    sheets: [{ sheet: 'Catalog', data: [{ sku: 'P-001', title: 'Desk lamp', price: 49.95 }] }]
  }];
  const buffer = await excelio.write(spec, { useWorker: true });
  assert.ok(Buffer.isBuffer(buffer));
  const result = await minified.read(buffer, { useWorker: true });
  assert.equal(result.ok, true);
  assert.equal(result.data[0].sheets[0].data[0].price, 49.95);
  const compactBuffer = await minified.write(spec, { useWorker: true });
  assert.ok(Buffer.isBuffer(compactBuffer));
  console.log('CommonJS and minified worker round-trips passed.');
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});