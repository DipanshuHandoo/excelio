import { fileURLToPath } from 'node:url';
import { run } from './run.mjs';

const root = fileURLToPath(new URL('../', import.meta.url));
run('npm', ['test'], { cwd: root });
run('npm', ['run', 'build'], { cwd: root });
run('npm', ['run', 'test:types'], { cwd: root });
run('npm', ['run', 'audit:dependencies'], { cwd: root });
run(process.execPath, ['scripts/check-package.mjs', '--built'], { cwd: root });
console.log('Release verification passed.');