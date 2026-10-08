/**
 * Custom error type for excelio. No external dependencies — keeps the package
 * portable to any Node project.
 *
 * Codes:
 *   INVALID_SPEC   — caller passed a malformed spec
 *   IO_ERROR       — corrupt input file, write failure, etc.
 *   ABORTED        — operation cancelled via AbortController
 *   WORKER_ERROR   — worker thread crashed or returned an error
 *   ROW_VALIDATION — a row failed validation (only thrown when onRowError === 'fail')
 */
export class ExcelIoError extends Error {
  constructor(message, { code = 'IO_ERROR', cause, details } = {}) {
    super(message);
    this.name = 'ExcelIoError';
    this.code = code;
    if (cause !== undefined) this.cause = cause;
    if (details !== undefined) this.details = details;
  }
}
