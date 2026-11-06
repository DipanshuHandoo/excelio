/**
 * Memory gate for iterable writes (release gate, not part of `verify`).
 * Run with: npm run test:memory   (node --expose-gc)
 *
 * The bound was measured against exceljs 4.4.0 / archiver 5.3.2; re-run on every exceljs upgrade.
 */
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { Writable } from 'node:stream';
import ExcelJS from 'exceljs';
import { excelio } from '../index.js';
import { __test } from '../writer.js';

// Measured peak RSS growth (scenario A: 50.3 MB on Windows, Node 22.23) plus 30 %; see README "Large exports".
const RSS_BUDGET_MB = Number(process.env.EXCELIO_RSS_BUDGET_MB || 65);
// Backlog is checked after each batch, so it may exceed the mark by one batch of row XML.
const BACKLOG_LIMIT = __test.OUTPUT_HIGH_WATER_MARK + 3 * 1024 * 1024;
const MB = 1024 * 1024;

const columns = ['id', 'url', 'target', 'title', 'occurrences', 'status'].map(key => ({ key }));
async function* rows(count, tag) {
  for (let i = 0; i < count; i++) {
    yield {
      id: `${tag}-${i}`,
      url: `https://example.com/page/${i}?q=${i}`,
      target: `https://example.org/${tag}/${i}/resource`,
      title: `Title number ${i} for ${tag}`,
      occurrences: i % 97,
      status: i % 3 === 0 ? 'ok' : 'redirect'
    };
  }
}

const slowSink = () => new Writable({
  highWaterMark: 64 * 1024,
  write(_chunk, _encoding, callback) { setTimeout(callback, 2); }
});

async function measure(name, fn) {
  global.gc?.();
  const baseline = process.memoryUsage().rss;
  let peak = baseline;
  __test.backlogStats.peak = 0;
  const sampler = setInterval(() => { peak = Math.max(peak, process.memoryUsage().rss); }, 250);
  const started = Date.now();
  try {
    await fn();
  } finally {
    clearInterval(sampler);
  }
  peak = Math.max(peak, process.memoryUsage().rss);
  const result = {
    name,
    seconds: ((Date.now() - started) / 1000).toFixed(1),
    peakRssGrowthMB: +((peak - baseline) / MB).toFixed(1),
    peakBacklogMB: +(__test.backlogStats.peak / MB).toFixed(1)
  };
  console.log(`  ${name}: ${result.seconds}s, peak RSS growth ${result.peakRssGrowthMB} MB, peak backlog ${result.peakBacklogMB} MB`);
  return result;
}

async function countDataRows(file) {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(file, { sharedStrings: 'cache', worksheets: 'emit' });
  let count = 0;
  for await (const worksheet of reader) {
    for await (const row of worksheet) if (row.number > 1) count++;
  }
  return count;
}

console.log(`excelio memory gate (budget ${RSS_BUDGET_MB} MB RSS growth)\n`);
const directory = await mkdtemp(path.join(tmpdir(), 'excelio-memory-'));
const failures = [];
const gate = (condition, message) => { if (!condition) failures.push(message); };

try {
  const fileA = path.join(directory, 'a.xlsx');
  const a = await measure('A. fast sink, 1,000,000 rows to file', () =>
    excelio.write([{ workbook: 'a', sheets: [{ sheet: 'Links', columns, data: rows(1_000_000, 'a') }] }], { to: fileA }));
  gate(a.peakRssGrowthMB < RSS_BUDGET_MB, `A: RSS growth ${a.peakRssGrowthMB} MB ≥ ${RSS_BUDGET_MB} MB`);
  const counted = await countDataRows(fileA);
  gate(counted === 1_000_000, `A: expected 1,000,000 data rows, counted ${counted}`);
  await rm(fileA, { force: true });

  const b = await measure('B. slow sink, 300,000 rows', () =>
    excelio.write([{ workbook: 'b', sheets: [{ sheet: 'Links', columns, data: rows(300_000, 'b') }] }], { to: slowSink() }));
  gate(b.peakRssGrowthMB < RSS_BUDGET_MB, `B: RSS growth ${b.peakRssGrowthMB} MB ≥ ${RSS_BUDGET_MB} MB`);
  gate(b.peakBacklogMB * MB <= BACKLOG_LIMIT, `B: backlog ${b.peakBacklogMB} MB above limit`);

  const c = await measure('C. slow sink, 2 × 200,000 rows', () =>
    excelio.write([{ workbook: 'c', sheets: [
      { sheet: 'One', columns, data: rows(200_000, 'c1') },
      { sheet: 'Two', columns, data: rows(200_000, 'c2') }
    ] }], { to: slowSink() }));
  gate(c.peakRssGrowthMB < RSS_BUDGET_MB, `C: RSS growth ${c.peakRssGrowthMB} MB ≥ ${RSS_BUDGET_MB} MB`);
  gate(c.peakBacklogMB * MB <= BACKLOG_LIMIT, `C: backlog ${c.peakBacklogMB} MB above limit`);

  const fileD = path.join(directory, 'd.xlsx');
  await measure('D. (info) scenario A with sharedStrings: true', () =>
    excelio.write([{ workbook: 'd', sheets: [{ sheet: 'Links', columns, data: rows(1_000_000, 'd') }] }], { to: fileD, sharedStrings: true }));
} finally {
  await rm(directory, { recursive: true, force: true });
}

assert.deepEqual(failures, [], failures.join('\n'));
console.log('\nMemory gate passed.');
