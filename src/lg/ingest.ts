/**
 * LG reconciliation — ingest a breakdown file into raw rows, then normalise
 * (GOAL.md §4 F1). Supports `.xlsx` (via exceljs) and `.csv`. Header validation
 * and row normalisation live in parse.ts; this layer only turns bytes into
 * `RawRow[]` and delegates.
 *
 * NOTE (scale, GOAL.md §5): this reads the whole workbook into memory via
 * exceljs `xlsx.load`. The ~550k-row production file works but is memory-heavy;
 * swapping to exceljs's streaming `WorkbookReader` is a follow-up (perf slice).
 */

import * as ExcelJS from 'exceljs';
import { ParseResult, RawCell, RawRow } from '../shared/models';
import { parseRows } from './parse';

export type IngestFormat = 'xlsx' | 'csv';

/** Unwraps an exceljs cell value (formula/richText/hyperlink objects) into a scalar. */
function unwrapCell(value: unknown): RawCell {
    if (value === null || value === undefined) {
        return undefined;
    }
    if (value instanceof Date) {
        return value;
    }
    if (typeof value === 'object') {
        const v = value as Record<string, unknown>;
        if ('result' in v) {
            return v.result as RawCell; // { formula, result }
        }
        if ('text' in v) {
            return v.text as RawCell; // { text, hyperlink }
        }
        if ('richText' in v && Array.isArray(v.richText)) {
            return (v.richText as { text?: string }[]).map((r) => r.text ?? '').join('');
        }
        if ('error' in v) {
            return undefined;
        }
        return String(value);
    }
    return value as RawCell;
}

/** Reads the first worksheet of an xlsx workbook into 0-indexed raw rows. */
export async function rowsFromXlsx(buffer: Buffer): Promise<RawRow[]> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer as unknown as ExcelJS.Buffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) {
        return [];
    }
    const rows: RawRow[] = [];
    worksheet.eachRow({ includeEmpty: false }, (row) => {
        // row.values is 1-indexed (index 0 is empty); rebuild a dense 0-indexed array.
        const values = row.values as unknown[];
        const out: RawRow = [];
        for (let c = 1; c < values.length; c++) {
            out[c - 1] = unwrapCell(values[c]);
        }
        rows.push(out);
    });
    return rows;
}

/** Minimal RFC-4180-ish CSV parser (quotes, escaped quotes, CRLF/LF). */
export function rowsFromCsv(input: string | Buffer): RawRow[] {
    const text = Buffer.isBuffer(input) ? input.toString('utf8') : input;
    const rows: RawRow[] = [];
    let row: string[] = [];
    let field = '';
    let inQuotes = false;
    let i = 0;
    const pushField = () => {
        row.push(field);
        field = '';
    };
    const pushRow = () => {
        pushField();
        rows.push(row as RawRow);
        row = [];
    };
    // Strip a UTF-8 BOM if present.
    if (text.charCodeAt(0) === 0xfeff) {
        i = 1;
    }
    for (; i < text.length; i++) {
        const ch = text[i];
        if (inQuotes) {
            if (ch === '"') {
                if (text[i + 1] === '"') {
                    field += '"';
                    i++;
                } else {
                    inQuotes = false;
                }
            } else {
                field += ch;
            }
        } else if (ch === '"') {
            inQuotes = true;
        } else if (ch === ',') {
            pushField();
        } else if (ch === '\r') {
            // handled by the \n branch; swallow lone \r before \n
            if (text[i + 1] === '\n') {
                i++;
            }
            pushRow();
        } else if (ch === '\n') {
            pushRow();
        } else {
            field += ch;
        }
    }
    // Flush the last field/row unless the input ended exactly on a newline.
    if (field !== '' || row.length > 0) {
        pushRow();
    }
    return rows;
}

/** Detects the file format from an explicit override, the filename, or the magic bytes. */
export function detectFormat(buffer: Buffer, filename?: string, override?: IngestFormat): IngestFormat {
    if (override) {
        return override;
    }
    if (filename) {
        const lower = filename.toLowerCase();
        if (lower.endsWith('.xlsx')) {
            return 'xlsx';
        }
        if (lower.endsWith('.csv')) {
            return 'csv';
        }
    }
    // xlsx is a zip archive → starts with "PK" (0x50 0x4B).
    if (buffer.length >= 2 && buffer[0] === 0x50 && buffer[1] === 0x4b) {
        return 'xlsx';
    }
    return 'csv';
}

export interface IngestOptions {
    filename?: string;
    format?: IngestFormat;
}

/** Ingests a breakdown file (xlsx or csv) into a ParseResult. */
export async function ingest(buffer: Buffer, options: IngestOptions = {}): Promise<ParseResult> {
    const format = detectFormat(buffer, options.filename, options.format);
    const rows = format === 'xlsx' ? await rowsFromXlsx(buffer) : rowsFromCsv(buffer);
    return parseRows(rows);
}
