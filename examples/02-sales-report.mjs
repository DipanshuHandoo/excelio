import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
import path from 'node:path';
import { excelio } from '@dipanshuhandoo/excelio';

const output = path.resolve(process.argv[2] ?? 'example-output');
await mkdir(output, { recursive: true });
const file = path.join(output, 'sales-report.xlsx');
const orders = [
  { orderId: 'ORD-1001', customer: 'Alice', amount: 1250.5, orderedAt: '2026-01-15' },
  { orderId: 'ORD-1002', customer: 'Bob', amount: 749.5, orderedAt: '2026-01-16' }
];
const spec = [{
  workbook: 'sales-2026',
  properties: { author: 'Sales team', title: 'January Sales', company: 'Example Company' },
  sheets: [
    {
      sheet: 'Orders/January',
      columns: [
        { key: 'orderId', header: 'Order ID', width: 18 },
        { key: 'customer', header: 'Customer', width: 24 },
        { key: 'amount', header: 'Total', type: 'number', format: '#,##0.00', width: 18 },
        { key: 'orderedAt', header: 'Order Date', type: 'date', format: 'yyyy-mm-dd', width: 18 }
      ],
      styling: { headerBold: true, freezeHeader: true, autoFilter: true },
      data: orders
    },
    {
      sheet: 'Summary',
      styling: { headerBold: true },
      data: [
        { metric: 'Order count', value: orders.length },
        { metric: 'Revenue', value: orders.reduce((total, order) => total + order.amount, 0) }
      ]
    }
  ]
}];

assert.equal(await excelio.write(spec, { to: file }), file);
const report = await excelio.read(file);
assert.equal(report.stats.totalSheets, 2);
assert.equal(report.data[0].sheets[0].sheet, 'Orders_January');
assert.equal(report.data[0].sheets[0].data[0]['Order ID'], 'ORD-1001');
assert.equal(report.data[0].sheets[0].data[0].Total, 1250.5);
assert.equal(report.data[0].sheets[1].data[1].value, 2000);
console.log(`Report saved: ${file}`);
console.log('Reading uses displayed headers: Order ID, Customer, Total, Order Date.');