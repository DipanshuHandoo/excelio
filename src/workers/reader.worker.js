import { parentPort, workerData } from 'node:worker_threads';
import { readSource } from '../reader.js';

(async () => {
  try {
    const { payload, options } = workerData;
    const result = await readSource(payload, options);
    parentPort.postMessage({ ok: true, result });
  } catch (err) {
    parentPort.postMessage({
      ok: false,
      error: { message: err.message, code: err.code, stack: err.stack, details: err.details }
    });
  }
})();
