import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { excelio } from '@dipanshuhandoo/excelio';

const output = path.resolve(process.argv[2] ?? 'example-output');
await mkdir(output, { recursive: true });
const inputFile = path.join(output, 'customers.xlsx');
const workbookJsonFile = path.join(output, 'customers-workbook.json');
const rowsJsonFile = path.join(output, 'customers-rows.json');

await excelio.write([{
  workbook: 'customer-export',
  sheets: [
    {
      sheet: 'Customers',
      columns: [
        { key: 'name' },
        { key: 'email' },
        { key: 'joined', type: 'date', format: 'yyyy-mm-dd' },
        { key: 'notes' }
      ],
      data: [
        { name: 'Alice', email: 'alice@example.com', joined: '2026-01-15', notes: null },
        { name: 'Bob', email: 'bob@example.com', joined: '2026-02-01', notes: 'Wholesale customer' }
      ]
    },
    {
      sheet: 'Summary',
      data: [{ metric: 'Customer count', value: 2 }]
    }
  ]
}], { to: inputFile });

const result = await excelio.read(inputFile);
assert.equal(result.ok, true);
assert.equal(result.stats.totalSheets, 2);
const firstSheetRows = result.data[0].sheets[0].data;
await writeFile(workbookJsonFile, JSON.stringify(result.data, null, 2) + '\n', 'utf8');
await writeFile(rowsJsonFile, JSON.stringify(firstSheetRows, null, 2) + '\n', 'utf8');

const workbookJson = JSON.parse(await readFile(workbookJsonFile, 'utf8'));
const rowsJson = JSON.parse(await readFile(rowsJsonFile, 'utf8'));
assert.deepEqual(workbookJson[0].sheets.map(sheet => sheet.sheet), ['Customers', 'Summary']);
assert.equal(workbookJson[0].sheets[1].data[0].value, 2);
assert.deepEqual(rowsJson, workbookJson[0].sheets[0].data);
assert.equal(rowsJson.length, 2);
assert.equal(rowsJson[0].joined, '2026-01-15T00:00:00.000Z');
assert.equal(rowsJson[0].notes, null);
assert.equal(Object.hasOwn(workbookJson[0], 'workbook'), false);

console.log(`Sample XLSX: ${inputFile}`);
console.log(`All-sheet JSON: ${workbookJsonFile}`);
console.log(`First-sheet rows JSON: ${rowsJsonFile}`);
console.log(JSON.stringify(rowsJson, null, 2));
console.log('Headers become JSON keys; dates are ISO strings and blank cells are null.');
console.log('JSON.stringify omits the undefined workbook label from the read result.');