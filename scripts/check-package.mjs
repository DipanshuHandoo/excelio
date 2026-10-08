import assert from 'node:assert/strict';
import { copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { run } from './run.mjs';

const root = fileURLToPath(new URL('../', import.meta.url));
const manifest = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
if (!process.argv.includes('--built')) run('npm', ['run', 'build'], { cwd: root });
const temporary = await mkdtemp(path.join(tmpdir(), 'excelio-package-'));

try {
  const packResult = JSON.parse(run('npm', ['pack', '--json', '--ignore-scripts', '--dry-run=false', '--pack-destination', temporary], { cwd: root, capture: true }));
  const packed = Array.isArray(packResult) ? packResult[0] : packResult[manifest.name];
  assert.ok(packed, 'npm pack did not return this package');
  assert.equal(packed.name, manifest.name);
  assert.equal(packed.version, manifest.version);
  const expected = new Set(['package.json', 'README.md', 'LICENSE', 'CHANGELOG.md']);
  for (const format of ['esm', 'cjs']) {
    for (const suffix of ['', '-min']) {
      const directory = `dist/${format}${suffix}`;
      const extension = format === 'cjs' ? 'cjs' : 'js';
      for (const entry of ['index', 'workers/reader.worker', 'workers/writer.worker']) {
        expected.add(`${directory}/${entry}.${extension}`);
        expected.add(`${directory}/${entry}.${extension}.map`);
      }
      expected.add(`${directory}/index.${format === 'cjs' ? 'd.cts' : 'd.ts'}`);
    }
  }
  const actual = new Set(packed.files.map(file => file.path));
  assert.deepEqual(actual, expected, 'Tarball contains missing or unexpected files');
  assert.ok(packed.files.every(file => file.size > 0), 'Tarball contains empty files');
  if (process.argv.includes('--list')) {
    console.log([...actual].sort().join('\n'));
    console.log(`Package contents verified: ${actual.size} files, ${packed.size} compressed bytes.`);
  } else {
    const consumer = path.join(temporary, 'consumer');
    await mkdir(path.join(consumer, 'types'), { recursive: true });
    await writeFile(path.join(consumer, 'package.json'), JSON.stringify({ name: 'excelio-installed-consumer', private: true, type: 'module' }));
    await copyFile(path.join(root, 'test/consumer.mjs'), path.join(consumer, 'consumer.mjs'));
    for (const extension of ['mts', 'cts']) {
      await copyFile(path.join(root, `test/types/consumer.${extension}`), path.join(consumer, `types/consumer.${extension}`));
    }
    const nodeTypes = JSON.parse(await readFile(path.join(root, 'node_modules/@types/node/package.json'), 'utf8'));
    run('npm', ['install', path.join(temporary, packed.filename), `@types/node@${nodeTypes.version}`,
      '--ignore-scripts', '--dry-run=false', '--no-audit', '--no-fund', '--package-lock=false'], { cwd: consumer });
    for (const format of ['esm', 'cjs', 'esm-min', 'cjs-min']) {
      run(process.execPath, ['consumer.mjs', format], { cwd: consumer });
    }
    run(process.execPath, [path.join(root, 'node_modules/typescript/bin/tsc'), '--noEmit', '--strict',
      '--module', 'nodenext', '--target', 'es2022', 'types/consumer.mts', 'types/consumer.cts'], { cwd: consumer });
    console.log('Installed tarball verified: readable/minified ESM/CommonJS, workers, streams, and declarations.');
  }
} finally {
  await rm(temporary, { recursive: true, force: true });
}