import { Worker } from 'node:worker_threads';
import { fileURLToPath } from 'node:url';
import { Buffer } from 'node:buffer';
import path from 'node:path';
import { statSync } from 'node:fs';

import { ExcelIoError } from './errors.js';

const DEFAULT_THRESHOLD = { rows: 50_000, bytes: 10 * 1024 * 1024 };

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const WORKER_PATHS = {
  write: path.join(__dirname, 'workers', 'writer.worker.js'),
  read:  path.join(__dirname, 'workers', 'reader.worker.js')
};

/**
 * Decide between inline execution and worker offload based on options.useWorker.
 *
 *   'auto'  → estimate workload; offload if it exceeds threshold
 *   true    → always offload
 *   false   → always inline
 *
 * Hooks (validateRow, beforeWrite, afterRead, onProgress, signal, logger) are
 * NOT serializable, so they can't cross the worker boundary. If any hook is
 * provided, we run inline regardless. Workers are best for "just transform a
 * spec/file" jobs — typical bulk imports/exports.
 */
export async function maybeUseWorker(op, payload, options, inlineFn) {
  const decision = decideMode(op, payload, options);
  if (decision.mode === 'inline') return inlineFn();

  return runInWorker(op, payload, sanitiseOptions(options));
}

function decideMode(op, payload, options) {
  const useWorker = options.useWorker ?? 'auto';

  if (hasHook(options)) return { mode: 'inline', reason: 'hooks present' };
  if (useWorker === false) return { mode: 'inline', reason: 'disabled' };
  if (useWorker === true)  return { mode: 'worker', reason: 'forced' };

  // auto: estimate
  const threshold = { ...DEFAULT_THRESHOLD, ...(options.workerThreshold || {}) };

  if (op === 'write') {
    const totalRows = estimateRows(payload);
    if (totalRows >= threshold.rows) return { mode: 'worker', reason: `rows ${totalRows} ≥ ${threshold.rows}` };
    return { mode: 'inline', reason: 'below threshold' };
  }

  if (op === 'read') {
    const bytes = estimateBytes(payload);
    if (bytes !== null && bytes >= threshold.bytes) return { mode: 'worker', reason: `bytes ${bytes} ≥ ${threshold.bytes}` };
    return { mode: 'inline', reason: 'below threshold or unknown size' };
  }

  return { mode: 'inline', reason: 'unknown op' };
}

function hasHook(options) {
  return Boolean(
    options.validateRow ||
    options.beforeWrite ||
    options.afterRead ||
    options.onProgress ||
    options.logger ||
    options.signal ||
    // streams can't cross thread boundaries via postMessage either
    (options.to && typeof options.to !== 'string')
  );
}

function estimateRows(spec) {
  if (!Array.isArray(spec)) return 0;
  let n = 0;
  for (const wb of spec) {
    for (const sheet of (wb.sheets || [])) {
      n += (sheet.data?.length || 0);
    }
  }
  return n;
}

function estimateBytes(input) {
  if (typeof input === 'string') {
    try { return statSync(input).size; } catch { return null; }
  }
  if (Buffer.isBuffer(input)) return input.length;
  return null; // streams: unknown
}

function sanitiseOptions(options) {
  // Drop anything that can't cross the postMessage boundary.
  const { validateRow, beforeWrite, afterRead, onProgress, signal, logger, to, ...rest } = options;
  void validateRow; void beforeWrite; void afterRead; void onProgress; void signal; void logger;
  return { ...rest, ...(typeof to === 'string' ? { to } : {}) };
}

function runInWorker(op, payload, options) {
  return new Promise((resolve, reject) => {
    const workerPath = WORKER_PATHS[op];
    if (!workerPath) {
      reject(new ExcelIoError(`No worker for op "${op}"`, { code: 'WORKER_ERROR' }));
      return;
    }

    const worker = new Worker(workerPath, {
      workerData: { op, payload: serialisePayload(payload), options }
    });

    worker.once('message', (msg) => {
      worker.terminate();
      if (msg.ok) resolve(msg.result);
      else reject(new ExcelIoError(msg.error?.message || 'Worker error', {
        code: msg.error?.code || 'WORKER_ERROR',
        cause: msg.error
      }));
    });
    worker.once('error', (err) => {
      worker.terminate();
      reject(new ExcelIoError(err.message, { code: 'WORKER_ERROR', cause: err }));
    });
    worker.once('exit', (code) => {
      if (code !== 0 && code !== null) {
        reject(new ExcelIoError(`Worker exited with code ${code}`, { code: 'WORKER_ERROR' }));
      }
    });
  });
}

function serialisePayload(payload) {
  // Buffers cross the boundary by reference (transferable). Strings are fine.
  // Spec arrays are JSON-serialisable. Streams should never reach here (gated above).
  return payload;
}
