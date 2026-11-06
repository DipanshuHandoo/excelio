import assert from 'node:assert/strict';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import { excelio } from '@dipanshuhandoo/excelio';

export async function jobsToXlsx(jobs) {
  const rows = jobs.map((job) => ({
    id: String(job._id),
    name: job.name || '',
    configName: job.configName || '',
    operation: job.operation || '',
    status: job.status || '',
    inserted: job.stats?.inserted ?? 0,
    updated: job.stats?.updated ?? 0,
    skipped: job.stats?.skipped ?? 0,
    errors: job.stats?.errors ?? 0,
    durationMs: job.durationMs ?? '',
    createdAt: job.createdAt ? new Date(job.createdAt).toISOString() : '',
    completedAt: job.completedAt ? new Date(job.completedAt).toISOString() : '',
  }));

  return excelio.write([
    {
      workbook: 'jobs',
      sheets: [
        {
          sheet: 'Jobs',
          columns: [
            { key: 'id', header: 'ID' },
            { key: 'name', header: 'Name' },
            { key: 'configName', header: 'Config' },
            { key: 'operation', header: 'Operation' },
            { key: 'status', header: 'Status' },
            { key: 'inserted', header: 'Inserted' },
            { key: 'updated', header: 'Updated' },
            { key: 'skipped', header: 'Skipped' },
            { key: 'errors', header: 'Errors' },
            { key: 'durationMs', header: 'Duration (ms)' },
            { key: 'createdAt', header: 'Created' },
            { key: 'completedAt', header: 'Completed' },
          ],
          data: rows,
        },
      ],
    },
  ]);
}

export async function manifestToXlsx(entries, jobName = 'Manifest') {
  const rows = entries.map((entry) => ({
    sequence: entry.sequence,
    collectionName: entry.collectionName,
    op: entry.op,
    docId: entry.docId != null ? String(entry.docId) : '',
    matchKey: entry.matchKey || '',
    matchValue: entry.matchValue != null ? String(entry.matchValue) : '',
  }));

  return excelio.write([
    {
      workbook: 'manifest',
      sheets: [
        {
          sheet: jobName.slice(0, 31),
          columns: [
            { key: 'sequence', header: 'Sequence' },
            { key: 'collectionName', header: 'Collection' },
            { key: 'op', header: 'Operation' },
            { key: 'docId', header: 'Doc ID' },
            { key: 'matchKey', header: 'Match Key' },
            { key: 'matchValue', header: 'Match Value' },
          ],
          data: rows,
        },
      ],
    },
  ]);
}

async function runExample() {
  const createdAt = '2026-10-09T10:00:00.000Z';
  const jobsBuffer = await jobsToXlsx([
    {
      _id: 'job-001',
      name: 'Customer clone',
      configName: 'staging',
      operation: 'clone',
      status: 'completed',
      stats: { inserted: 12, updated: 3 },
      durationMs: 1500,
      createdAt,
    },
    { _id: 'job-002' },
  ]);
  assert.ok(Buffer.isBuffer(jobsBuffer));
  const jobsWorkbook = new ExcelJS.Workbook();
  await jobsWorkbook.xlsx.load(jobsBuffer);
  assert.equal(jobsWorkbook.worksheets.length, 1);
  const jobsSheet = jobsWorkbook.getWorksheet('Jobs');
  assert.deepEqual(jobsSheet.getRow(1).values.slice(1), [
    'ID', 'Name', 'Config', 'Operation', 'Status', 'Inserted',
    'Updated', 'Skipped', 'Errors', 'Duration (ms)', 'Created', 'Completed',
  ]);
  assert.deepEqual(jobsSheet.getRow(2).values.slice(1), [
    'job-001', 'Customer clone', 'staging', 'clone', 'completed',
    12, 3, 0, 0, 1500, createdAt, '',
  ]);
  assert.deepEqual(jobsSheet.getRow(3).values.slice(1), [
    'job-002', '', '', '', '', 0, 0, 0, 0, '', '', '',
  ]);

  const jobName = 'Customer clone manifest with a long name';
  const manifestBuffer = await manifestToXlsx([
    { sequence: 1, collectionName: 'customers', op: 'insert', docId: 123, matchValue: 0 },
  ], jobName);
  assert.ok(Buffer.isBuffer(manifestBuffer));
  const manifestWorkbook = new ExcelJS.Workbook();
  await manifestWorkbook.xlsx.load(manifestBuffer);
  assert.equal(manifestWorkbook.worksheets.length, 1);
  const manifestSheet = manifestWorkbook.getWorksheet(jobName.slice(0, 31));
  assert.deepEqual(manifestSheet.getRow(1).values.slice(1), [
    'Sequence', 'Collection', 'Operation', 'Doc ID', 'Match Key', 'Match Value',
  ]);
  assert.deepEqual(manifestSheet.getRow(2).values.slice(1), [
    1, 'customers', 'insert', '123', '', '0',
  ]);

  const emptyBuffer = await manifestToXlsx([]);
  const emptyWorkbook = new ExcelJS.Workbook();
  await emptyWorkbook.xlsx.load(emptyBuffer);
  assert.equal(emptyWorkbook.getWorksheet('Manifest').rowCount, 1);
  console.log('Service exports verified: XLSX buffers, headers, rows, defaults, and sheet names.');
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await runExample();
}