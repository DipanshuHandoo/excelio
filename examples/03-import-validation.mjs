import assert from 'node:assert/strict';
import { excelio, ExcelIoError } from '@dipanshuhandoo/excelio';

const uploadedWorkbook = await excelio.write([{
  workbook: 'customer-upload',
  sheets: [{
    sheet: 'Customers',
    data: [
      { name: 'Alice', email: ' ALICE@EXAMPLE.COM ', credit: 100 },
      { name: 'Bob', email: 'not-an-email', credit: 50 },
      { name: 'Charlie', email: 'charlie@example.com', credit: -10 }
    ]
  }]
}]);

const importOptions = {
  afterRead: row => ({ ...row, email: String(row.email ?? '').trim().toLowerCase() }),
  validateRow(row) {
    if (!row.email.includes('@')) {
      return { code: 'INVALID_EMAIL', column: 'email', message: 'Email must contain @' };
    }
    if (row.credit < 0) {
      return { code: 'NEGATIVE_CREDIT', column: 'credit', message: 'Credit cannot be negative' };
    }
    return null;
  }
};

const result = await excelio.read(uploadedWorkbook, { ...importOptions, onRowError: 'collect' });
assert.equal(result.ok, false);
assert.equal(result.stats.totalRows, 1);
assert.equal(result.errors.length, 2);
assert.equal(result.data[0].sheets[0].data[0].email, 'alice@example.com');
console.log('Accepted customers:', result.data[0].sheets[0].data);
console.log('Rejected rows:', result.errors);

try {
  await excelio.read(uploadedWorkbook, { ...importOptions, onRowError: 'fail' });
  assert.fail('Strict import should reject invalid rows');
} catch (error) {
  if (!(error instanceof ExcelIoError) || error.code !== 'ROW_VALIDATION') throw error;
  assert.equal(error.details.row, 3);
  console.log(`Strict import stopped at Excel row ${error.details.row}: ${error.message}`);
}