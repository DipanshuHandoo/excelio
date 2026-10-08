import { spawnSync } from 'node:child_process';

export function run(command, args, { cwd, capture = false } = {}) {
  if (command === 'npm') {
    if (!process.env.npm_execpath) throw new Error('Run this check through its npm script.');
    args = [process.env.npm_execpath, ...args];
    command = process.execPath;
  }
  const result = spawnSync(command, args, {
    cwd,
    encoding: 'utf8',
    stdio: capture ? 'pipe' : 'inherit'
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`Command failed (${result.status}): ${command} ${args.join(' ')}\n${result.stdout ?? ''}${result.stderr ?? ''}`);
  }
  return result.stdout;
}