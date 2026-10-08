import { build } from 'esbuild';
import { copyFile, mkdir, readFile, rm, stat } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const root = fileURLToPath(new URL('../', import.meta.url));
const dist = path.join(root, 'dist');
const minOnly = process.argv.includes('--min');
if (!minOnly) await rm(dist, { recursive: true, force: true });

for (const format of ['esm', 'cjs']) {
  for (const minify of minOnly ? [true] : [false, true]) {
    const extension = format === 'cjs' ? 'cjs' : 'js';
    const directory = path.join(dist, `${format}${minify ? '-min' : ''}`);
    await rm(directory, { recursive: true, force: true });
    await mkdir(directory, { recursive: true });
    await build({
      absWorkingDir: root,
      entryPoints: {
        index: 'src/index.js',
        'workers/reader.worker': 'src/workers/reader.worker.js',
        'workers/writer.worker': 'src/workers/writer.worker.js'
      },
      outdir: directory,
      outExtension: { '.js': `.${extension}` },
      bundle: true,
      packages: 'external',
      platform: 'node',
      format,
      target: 'node22',
      minify,
      sourcemap: true,
      legalComments: 'none',
      define: {
        __EXCELIO_WORKER_EXT__: JSON.stringify(`.${extension}`),
        ...(format === 'cjs' ? { __EXCELIO_MODULE_DIR__: '__dirname', 'import.meta.url': 'undefined' } : {})
      },
      logLevel: 'warning'
    });
    await copyFile(path.join(root, 'src/excelio.d.ts'), path.join(directory, format === 'cjs' ? 'index.d.cts' : 'index.d.ts'));
    for (const filename of [`index.${extension}`, `index.${extension}.map`,
      `workers/reader.worker.${extension}`, `workers/writer.worker.${extension}`,
      format === 'cjs' ? 'index.d.cts' : 'index.d.ts']) {
      if (!(await stat(path.join(directory, filename))).size) throw new Error(`Empty artifact: ${filename}`);
    }
  }
}

if (!minOnly) {
  const manifest = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  for (const entry of [manifest.exports['.'], manifest.exports['./min']]) {
    for (const condition of ['import', 'require']) {
      await stat(path.join(root, entry[condition].types));
      await stat(path.join(root, entry[condition].default));
    }
  }
}
console.log(`Built ${minOnly ? 'minified' : 'readable and minified'} ESM/CommonJS, workers, source maps, and declarations.`);