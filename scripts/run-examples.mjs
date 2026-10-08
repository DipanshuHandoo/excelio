import { mkdtemp, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { run } from './run.mjs';

const root = fileURLToPath(new URL('../', import.meta.url));
const available = (await readdir(path.join(root, 'examples')))
  .filter(filename => /^\d\d-[a-z-]+\.(mjs|cjs)$/.test(filename))
  .sort();
const requested = process.argv.slice(2).filter(argument => argument !== '--built');
const selected = requested.length
  ? available.filter(filename => requested.includes(filename) || requested.includes(path.parse(filename).name))
  : available;
for (const argument of requested) {
  if (!available.some(filename => filename === argument || path.parse(filename).name === argument)) {
    throw new Error(`Unknown example: ${argument}. Choose from: ${available.join(', ')}`);
  }
}
if (!selected.length) throw new Error('No runnable examples found.');
if (!process.argv.includes('--built')) run('npm', ['run', 'build'], { cwd: root });
const output = await mkdtemp(path.join(tmpdir(), 'excelio-examples-'));
try {
  for (const filename of selected) {
    console.log(`\nExample: ${filename}`);
    run(process.execPath, [path.join(root, 'examples', filename), output], { cwd: root });
  }
} finally {
  await rm(output, { recursive: true, force: true });
}
console.log(`\n${selected.length} example${selected.length === 1 ? '' : 's'} passed. Temporary output cleaned up.`);