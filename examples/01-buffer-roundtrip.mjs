import assert from 'node:assert/strict';
import { Buffer } from 'node:buffer';
import { excelio } from '@dipanshuhandoo/excelio';

const spec = [{
  workbook: 'employees',
  sheets: [{
    sheet: 'Employees',
    columns: [
      { key: 'name', width: 24 },
      { key: 'salary', type: 'number', format: '#,##0.00' },
      { key: 'active', type: 'boolean' },
      { key: 'joined', type: 'date' },
      { key: 'notes', width: 30 }
    ],
    data: [
      { name: 'Alice', salary: '72000.50', active: 'true', joined: '2026-01-15', notes: 'Remote\nFull-time' },
      { name: 'Bob', salary: 65000, active: false, joined: new Date('2026-02-01'), notes: null }
    ]
  }]
}];

const buffer = await excelio.write(spec);
assert.ok(Buffer.isBuffer(buffer));
const result = await excelio.read(buffer);
assert.equal(result.ok, true);
assert.equal(result.stats.totalRows, 2);
const employees = result.data[0].sheets[0].data;
assert.equal(employees[0].salary, 72000.5);
assert.equal(employees[0].active, true);
assert.equal(employees[0].joined, '2026-01-15T00:00:00.000Z');
assert.equal(employees[0].notes, 'Remote\nFull-time');
assert.equal(employees[1].notes, null);

const epochs = await excelio.read(buffer, { dateFormat: 'epoch' });
assert.equal(epochs.data[0].sheets[0].data[0].joined, Date.parse('2026-01-15'));
console.log(JSON.stringify({ bytes: buffer.length, employees }, null, 2));