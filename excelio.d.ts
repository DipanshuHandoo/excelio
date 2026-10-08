/**
 * excelio — Plug-and-Play Excel ↔ JSON utility.
 * TypeScript declarations for consumers; the package itself is plain JS.
 */
import type { Readable, Writable } from 'node:stream';

export type ExcelIoErrorCode =
  | 'INVALID_SPEC'
  | 'IO_ERROR'
  | 'ABORTED'
  | 'WORKER_ERROR'
  | 'ROW_VALIDATION';

export class ExcelIoError extends Error {
  code: ExcelIoErrorCode;
  cause?: unknown;
  details?: unknown;
}

export type ColumnType = 'string' | 'number' | 'boolean' | 'date';

export interface ColumnSpec {
  /** Property name on each data row. */
  key: string;
  /** Header text in row 1. Defaults to `key`. */
  header?: string;
  /** Type hint that informs cell formatting and write-side coercion. */
  type?: ColumnType;
  /** Excel number-format string (e.g. 'yyyy-mm-dd', '#,##0.00'). */
  format?: string;
  /** Column width in Excel units. */
  width?: number;
}

export interface SheetStyling {
  headerBold?: boolean;
  freezeHeader?: boolean;
  autoFilter?: boolean;
}

export interface SheetSpec<Row = Record<string, unknown>> {
  sheet: string;
  columns?: ColumnSpec[];
  styling?: SheetStyling;
  data: Row[];
}

export interface WorkbookSpec<Row = Record<string, unknown>> {
  workbook: string;
  properties?: {
    author?: string;
    company?: string;
    title?: string;
    created?: Date | string;
    modified?: Date | string;
  };
  sheets: SheetSpec<Row>[];
}

export type Spec<Row = Record<string, unknown>> = WorkbookSpec<Row>[];

export interface RowContext {
  workbookIndex: number;
  workbookName?: string;
  sheetIndex: number;
  sheetName: string;
  rowNumber: number;
}

export interface RowError {
  workbook: number;
  sheet: string;
  row: number;
  column: string | null;
  code: string;
  message: string;
}

export interface ReadResult<Row = Record<string, unknown>> {
  ok: boolean;
  data: Spec<Row>;
  errors: RowError[];
  stats: {
    totalRows: number;
    totalSheets: number;
    durationMs: number;
  };
}

export interface ProgressEvent {
  phase: 'read' | 'write';
  workbookIndex?: number;
  sheetIndex?: number;
  sheetName?: string;
  rowsProcessed: number;
  totalRows?: number;
}

export interface Logger {
  info?: (msg: string, ...args: unknown[]) => void;
  warn?: (msg: string, ...args: unknown[]) => void;
  error?: (msg: string, ...args: unknown[]) => void;
}

export interface CommonOptions {
  /** Called per row at every batchSize boundary. */
  onProgress?: (e: ProgressEvent) => void;
  /** AbortController signal; checked at every batch boundary. */
  signal?: AbortSignal;
  /** BYO logger. Default: no-op. */
  logger?: Logger;
  /** Hooks force inline execution (workers can't serialise functions). */
  batchSize?: number;
  /** 'auto' (default), true (always), false (never). */
  useWorker?: 'auto' | boolean;
  /** Auto-mode threshold. Default: { rows: 50_000, bytes: 10MB }. */
  workerThreshold?: { rows?: number; bytes?: number };
}

export interface WriteOptions<Row = Record<string, unknown>> extends CommonOptions {
  /** File path or Writable stream. Omit to receive Buffer(s) back. */
  to?: string | Writable;
  /** Per-row transform applied before writing. */
  beforeWrite?: (row: Row, ctx: RowContext) => Row | void;
}

export interface ReadOptions<Row = Record<string, unknown>> extends CommonOptions {
  /** Row number containing headers (1-indexed). Default: 1. */
  headerRow?: number;
  /** Empty cells become null instead of ''. Default: true. */
  nullForBlank?: boolean;
  /** Trim whitespace from string cells. Default: true. */
  trimStrings?: boolean;
  /** Date output format. Default: 'iso'. */
  dateFormat?: 'iso' | 'date' | 'epoch';
  /** Per-row transform applied after parsing, before validation. */
  afterRead?: (row: Row, ctx: RowContext) => Row | void;
  /** Per-row validation. Return null/undefined for valid; an issue object for invalid. */
  validateRow?: (row: Row, ctx: RowContext) => null | undefined | { code?: string; message: string; column?: string };
  /** 'collect' (default) returns errors in the result; 'fail' throws on first. */
  onRowError?: 'collect' | 'fail';
}

export interface Excelio {
  write<Row = Record<string, unknown>>(spec: Spec<Row>, options?: WriteOptions<Row>): Promise<Buffer | Buffer[] | string | Writable>;
  writeStream<Row = Record<string, unknown>>(spec: Spec<Row>, options?: WriteOptions<Row>): Readable;
  read<Row = Record<string, unknown>>(input: Buffer | string | Readable, options?: ReadOptions<Row>): Promise<ReadResult<Row>>;
}

export const excelio: Excelio;
export default excelio;
