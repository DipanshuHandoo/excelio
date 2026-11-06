/// <reference types="node" />
/**
 * excelio — Plug-and-Play Excel ↔ JSON utility.
 * TypeScript declarations for consumers; the package itself is plain JS.
 */
import type { Readable, Writable } from 'node:stream';
import type { Buffer } from 'node:buffer';

export type ExcelIoErrorCode =
  | 'INVALID_SPEC'
  | 'IO_ERROR'
  | 'ABORTED'
  | 'WORKER_ERROR'
  | 'ROW_VALIDATION'
  | 'SOURCE_ERROR';

export class ExcelIoError extends Error {
  constructor(message: string, options?: { code?: ExcelIoErrorCode; cause?: unknown; details?: unknown });
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

/**
 * Rows for one sheet: an array, or any sync/async iterable (generator, database cursor,
 * object-mode stream). Iterables are consumed once and always written inline.
 */
export type RowSource<Row> = readonly Row[] | Iterable<Row> | AsyncIterable<Row>;

export interface SheetSpec<Row = Record<string, unknown>> {
  sheet: string;
  columns?: ColumnSpec[];
  styling?: SheetStyling;
  /** Omitted or null means an empty sheet. */
  data?: RowSource<Row> | null;
  /** Optional row-count hint used for progress `totalRows` when `data` is an iterable. */
  totalRows?: number;
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

export interface ParsedColumn {
  key: string;
  header: string;
}

export interface ParsedSheet<Row = Record<string, unknown>> {
  sheet: string;
  columns: ParsedColumn[];
  data: Row[];
}

export interface ParsedWorkbook<Row = Record<string, unknown>> {
  workbook: undefined;
  sheets: ParsedSheet<Row>[];
}

export interface RowContext {
  workbookIndex: number;
  workbookName?: string;
  sheetIndex: number;
  sheetName: string;
  /** 1 for the first worksheet of a logical sheet, 2+ for continuation sheets. Write only. */
  sheetPart?: number;
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
  data: ParsedWorkbook<Row>[];
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
  /** Write only: worksheet part of the logical sheet (see `maxRowsPerSheet`). */
  sheetPart?: number;
  rowsProcessed: number;
  /** Undefined when an iterable sheet has no `totalRows` hint. */
  totalRows?: number;
}

export interface Logger {
  info?: (msg: string, ...args: unknown[]) => void;
  warn?: (msg: string, ...args: unknown[]) => void;
  error?: (msg: string, ...args: unknown[]) => void;
}

export interface CommonOptions {
  /** Called at batchSize row boundaries; writing also emits a final update. */
  onProgress?: (e: ProgressEvent) => void;
  /** Checked between rows/sheets; iterable sheets also observe it while waiting for output capacity. */
  signal?: AbortSignal;
  /** BYO logger. Default: no-op. */
  logger?: Logger;
  /** Positive integer. Progress interval; iterable sheets also yield and bound output backlog at this interval. */
  batchSize?: number;
  /** 'auto' (default), true (request worker), false (inline). Hooks/streams force inline. */
  useWorker?: 'auto' | boolean;
  /** Auto-mode threshold. Default: { rows: 50_000, bytes: 10MB }. */
  workerThreshold?: { rows?: number; bytes?: number };
}

export interface WriteOptions<Row = Record<string, unknown>> extends CommonOptions {
  /** File path or Writable stream. Omit to receive Buffer(s) back. */
  to?: string | Writable;
  /** Per-row transform applied before writing. */
  beforeWrite?: (row: Row, ctx: RowContext) => Row | void;
  /** Data rows per worksheet before continuing on "<name> (2)". Default and maximum 1,048,575. */
  maxRowsPerSheet?: number;
  /** Excel shared-strings table. 'auto' (default): on unless any sheet has iterable data. */
  sharedStrings?: 'auto' | boolean;
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
  write<Row = Record<string, unknown>, Target extends string | Writable = string | Writable>(spec: Spec<Row>, options: WriteOptions<Row> & { to: Target }): Promise<Target>;
  write<Row = Record<string, unknown>>(spec: Spec<Row>, options?: WriteOptions<Row> & { to?: undefined }): Promise<Buffer | Buffer[]>;
  write<Row = Record<string, unknown>>(spec: Spec<Row>, options?: WriteOptions<Row>): Promise<Buffer | Buffer[] | string | Writable>;
  writeStream<Row = Record<string, unknown>>(spec: Spec<Row>, options?: WriteOptions<Row>): Readable;
  read<Row = Record<string, unknown>>(input: Buffer | string | Readable, options?: ReadOptions<Row>): Promise<ReadResult<Row>>;
}

export const excelio: Excelio;
export default excelio;
