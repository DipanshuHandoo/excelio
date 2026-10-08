import { parentPort, workerData } from 'node:worker_threads';
import { writeSpec } from '../writer.js';

(async () => {
  try {
    const { payload, options } = workerData;
    const result = await writeSpec(payload, options);
    parentPort.postMessage({ ok: true, result });
  } catch (err) {
    parentPort.postMessage({
      ok: false,
      error: { message: err.message, code: err.code, stack: err.stack, details: err.details }
    });
  }
})();
