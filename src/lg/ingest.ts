/**
 * LG reconciliation — ingest a breakdown file into raw rows, then normalise
 * (GOAL.md §4 F1). Supports `.xlsx` (via the exceljs *streaming* WorkbookReader —
 * rows are read one at a time instead of materialising the whole workbook DOM,
 * which matters at the ~550k-row production size, GOAL.md §5) and `.csv`.
 *
 * Every worksheet is read (GOAL.md §2.3): sheets whose header matches the breakdown
 * schema contribute rows; others are skipped with a SHEET_SKIPPED note. Legacy `.xls`
 * (BIFF/OLE) files are rejected with UNSUPPORTED_FORMAT — exceljs cannot read them.
 * Header validation and row normalisation live in parse.ts; this layer only turns
 * bytes into rows and delegates.
 */

import { Readable } from 'node:stream';
import * as ExcelJS from 'exceljs';
import { GL_CATALOG, LgRunMode, ParseError, ParsedPosting, ParseResult, RawCell, RawRow, RegisterCheque, resolveGlCode } from '../shared/models';
import { detectSheetRole, resolveMode } from './detect';
import { parseSheets, SheetRows } from './parse';
import { parseRegisterSheet } from './registerParse';
import { parseStatementSheet } from './statementParse';

export type IngestFormat = 'xlsx' | 'csv';

/** ParseResult plus the input family and (register mode) the parsed cheque register. */
export type IngestResult = ParseResult & {
    mode: LgRunMode;
    cheques?: RegisterCheque[];
};

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
            const result = v.result; // { formula, result }
            if (result !== null && typeof result === 'object' && !(result instanceof Date)) {
                return undefined; // cached formula ERROR result, e.g. { error: '#N/A' }
            }
            return result as RawCell;
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

/** OLE compound-file magic bytes — a legacy `.xls` (BIFF) workbook, which exceljs cannot read. */
export function isLegacyXls(buffer: Buffer): boolean {
    const magic = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
    return buffer.length >= magic.length && magic.every((byte, i) => buffer[i] === byte);
}

/**
 * Streams every worksheet of an xlsx workbook into raw 0-indexed rows. Uses the
 * exceljs streaming WorkbookReader: rows are emitted one at a time, so memory holds
 * raw cell values only — never the full workbook object model.
 */
async function sheetsFromXlsxStream(buffer: Buffer): Promise<SheetRows[]> {
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from([buffer]), {
        entries: 'emit',
        sharedStrings: 'cache',
        hyperlinks: 'ignore',
        styles: 'cache',
        worksheets: 'emit',
    });
    const sheets: SheetRows[] = [];
    let index = 0;
    for await (const worksheet of reader) {
        index++;
        const rows: RawRow[] = [];
        for await (const row of worksheet) {
            // row.values is 1-indexed (index 0 is empty); rebuild a dense 0-indexed array.
            const values = row.values as unknown[];
            const out: RawRow = [];
            for (let c = 1; c < values.length; c++) {
                out[c - 1] = unwrapCell(values[c]);
            }
            rows.push(out);
        }
        const name = (worksheet as unknown as { name?: string }).name;
        sheets.push({ name: name || `Sheet${index}`, rows });
    }
    return sheets;
}

/** Non-streaming fallback: the full workbook DOM, reading every worksheet. */
async function sheetsFromXlsxDom(buffer: Buffer): Promise<SheetRows[]> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer as unknown as ExcelJS.Buffer);
    return workbook.worksheets.map((worksheet, i) => {
        const rows: RawRow[] = [];
        worksheet.eachRow({ includeEmpty: false }, (row) => {
            const values = row.values as unknown[];
            const out: RawRow = [];
            for (let c = 1; c < values.length; c++) {
                out[c - 1] = unwrapCell(values[c]);
            }
            rows.push(out);
        });
        return { name: worksheet.name || `Sheet${i + 1}`, rows };
    });
}

/**
 * Reads every worksheet, preferring the streaming reader (GOAL.md §5 scale). The
 * exceljs streaming reader throws when a worksheet zip entry precedes xl/workbook.xml
 * (entry order is not guaranteed by all writers), so fall back to the in-memory
 * reader rather than failing the upload.
 */
export async function sheetsFromXlsx(buffer: Buffer): Promise<SheetRows[]> {
    try {
        return await sheetsFromXlsxStream(buffer);
    } catch {
        return sheetsFromXlsxDom(buffer);
    }
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

function emptyResult(errors: ParseError[]): ParseResult {
    return {
        postings: [],
        errors,
        summary: { dataRows: 0, parsed: 0, debitCount: 0, creditCount: 0, netFils: 0, currencies: [], branches: [] },
    };
}

/**
 * Parses a register-family workbook (GOAL-3 §4.1): ledger-statement sheets feed
 * postings (continuous row numbering across sheets), register sheets feed
 * cheques, anything else is SHEET_SKIPPED. Direction is per row, so the
 * Credit/Debit sheet split is irrelevant to correctness.
 */
function parseRegisterFamily(sheets: SheetRows[], roles: ReturnType<typeof detectSheetRole>[]): IngestResult {
    const postings: ParsedPosting[] = [];
    const cheques: RegisterCheque[] = [];
    const errors: ParseError[] = [];
    let ledgerRows = 0;
    let registerRows = 0;

    for (let i = 0; i < sheets.length; i++) {
        const sheet = sheets[i];
        if (roles[i] === 'ledgerStatement') {
            const result = parseStatementSheet(sheet.rows, sheet.name, ledgerRows);
            postings.push(...result.postings);
            errors.push(...result.errors);
            ledgerRows += result.dataRows;
        } else if (roles[i] === 'register') {
            const result = parseRegisterSheet(sheet.rows, sheet.name);
            // Keep cheque row numbers unique across multiple register sheets.
            cheques.push(...result.cheques.map((c) => ({ ...c, rowNumber: c.rowNumber + registerRows })));
            errors.push(...result.errors);
            registerRows += result.dataRows;
        } else {
            errors.push({
                code: 'SHEET_SKIPPED',
                sheet: sheet.name,
                message: `Worksheet ${
                    sheet.name ? `"${sheet.name}"` : '(unnamed)'
                } matches neither the ledger-statement nor the register schema — skipped`,
            });
        }
    }

    let debitCount = 0;
    let creditCount = 0;
    let netFils = 0;
    const currencies = new Set<string>();
    const branches = new Set<string>();
    for (const p of postings) {
        netFils += p.amountBhdFils;
        if (p.direction === 'debit') {
            debitCount++;
        } else {
            creditCount++;
        }
        currencies.add(p.currency);
        branches.add(p.branchNumber);
    }

    return {
        mode: 'register',
        postings,
        cheques,
        errors,
        summary: {
            dataRows: ledgerRows + registerRows,
            parsed: postings.length,
            debitCount,
            creditCount,
            netFils,
            currencies: [...currencies].sort(),
            branches: [...branches].sort(),
        },
    };
}

/** Parses a statement-only workbook (GOAL-8): every ledgerStatement sheet feeds
 *  postings via parseStatementSheet (continuous row numbers); no register, no cheques. */
function parseStatementFamily(sheets: SheetRows[], roles: ReturnType<typeof detectSheetRole>[]): IngestResult {
    const postings: ParsedPosting[] = [];
    const errors: ParseError[] = [];
    let ledgerRows = 0;
    for (let i = 0; i < sheets.length; i++) {
        if (roles[i] === 'ledgerStatement') {
            const result = parseStatementSheet(sheets[i].rows, sheets[i].name, ledgerRows);
            postings.push(...result.postings);
            errors.push(...result.errors);
            ledgerRows += result.dataRows;
        } else {
            errors.push({
                code: 'SHEET_SKIPPED',
                sheet: sheets[i].name,
                message: `Worksheet ${sheets[i].name ? `"${sheets[i].name}"` : '(unnamed)'} is not a ledger-statement sheet — skipped`,
            });
        }
    }
    let debitCount = 0;
    let creditCount = 0;
    let netFils = 0;
    const currencies = new Set<string>();
    const branches = new Set<string>();
    for (const p of postings) {
        netFils += p.amountBhdFils;
        if (p.direction === 'debit') {
            debitCount++;
        } else {
            creditCount++;
        }
        currencies.add(p.currency);
        branches.add(p.branchNumber);
    }
    return {
        mode: 'statement',
        postings,
        errors,
        summary: {
            dataRows: ledgerRows,
            parsed: postings.length,
            debitCount,
            creditCount,
            netFils,
            currencies: [...currencies].sort(),
            branches: [...branches].sort(),
        },
    };
}

/** The embedded GL of a statement-only workbook = the resolved Nostro/BGL Account
 *  of its first parsed posting. undefined when nothing resolves. */
function statementGlOf(postings: ParsedPosting[]): ReturnType<typeof resolveGlCode> {
    for (const p of postings) {
        const code = resolveGlCode(p.gl);
        if (code) {
            return code;
        }
    }
    return undefined;
}

/**
 * Classifies pooled worksheets by role (GOAL-3 §4.1) and parses them:
 * breakdown sheets keep the original pipeline; register-family sheets (ledger
 * statements + cheque register) parse into postings AND cheques; mixed or
 * half-register inputs are rejected.
 */
function ingestSheets(sheets: SheetRows[]): IngestResult {
    const roles = sheets.map((sheet) => detectSheetRole(sheet.rows));

    // GOAL-8: a statement-only workbook is a standalone running-balance statement
    // IFF its embedded account resolves to a statement-mode GL. Otherwise it keeps
    // the register-completeness rule (INCOMPLETE_REGISTER_INPUT) — a lone MGR ledger
    // extract still means "the register file is missing".
    const statementOnly =
        roles.includes('ledgerStatement') && !roles.includes('register') && !roles.includes('breakdown');
    if (statementOnly) {
        const statement = parseStatementFamily(sheets, roles);
        const code = statementGlOf(statement.postings);
        if (code && GL_CATALOG[code].mode === 'statement') {
            return statement;
        }
        const { error } = resolveMode(roles); // the exact existing INCOMPLETE_REGISTER_INPUT message
        return { mode: 'breakdown', ...emptyResult([error!]) };
    }

    const { mode, error } = resolveMode(roles);
    if (error) {
        return { mode: 'breakdown', ...emptyResult([error]) };
    }
    if (mode === 'register') {
        return parseRegisterFamily(sheets, roles);
    }
    // Breakdown mode — and the nothing-recognisable case, which parseSheets
    // reports exactly as before (MISSING_HEADER details from the first sheet).
    return { mode: 'breakdown', ...parseSheets(sheets) };
}

/**
 * Ingests an uploaded file (xlsx or csv, any number of worksheets). Worksheets
 * are classified by role (GOAL-3 §4.1): breakdown workbooks keep the original
 * pipeline; register-family workbooks (ledger statements + cheque register)
 * parse into postings AND cheques; mixed or half-register inputs are rejected.
 */
export async function ingest(buffer: Buffer, options: IngestOptions = {}): Promise<IngestResult> {
    if (isLegacyXls(buffer)) {
        return {
            mode: 'breakdown',
            ...emptyResult([
                {
                    code: 'UNSUPPORTED_FORMAT',
                    message:
                        'Legacy .xls (BIFF) and password-protected workbooks are not supported — re-save the file as an unencrypted .xlsx or export .csv',
                },
            ]),
        };
    }
    const format = detectFormat(buffer, options.filename, options.format);
    if (format === 'csv') {
        // GOAL-8: route CSV through role detection too, so a statement-schema CSV
        // (e.g. a sanitised single-sheet VAT file) reaches the statement path.
        // A breakdown CSV still detects role 'breakdown' → byte-identical result.
        // Keep the sheet NAMELESS (as the old `parseSheets([{ rows }])` call did) so
        // breakdown-CSV postings keep `sheet: undefined` and sheetBalances keeps its
        // '(sheet)' fallback label — no behaviour change for existing breakdown CSVs.
        return ingestSheets([{ rows: rowsFromCsv(buffer) }]);
    }
    return ingestSheets(await sheetsFromXlsx(buffer));
}

/** One file of a (possibly multi-file) upload. */
export interface UploadFile {
    buffer: Buffer;
    filename?: string;
    format?: IngestFormat;
}

/**
 * Ingests one or more uploaded files as a SINGLE run: every file's worksheets
 * pool together BEFORE role detection, so a register family split across files
 * (the ledger extract in one, the cheque register in another) resolves — while
 * mixed families still reject with MIXED_MODE. In multi-file uploads each
 * sheet name gains a "filename › sheet" prefix (csv contributes one sheet
 * named after its file) so errors and postings keep per-file provenance.
 * A single file goes through ingest() unchanged — byte-identical behaviour.
 */
export async function ingestFiles(files: UploadFile[]): Promise<IngestResult> {
    if (files.length === 0) {
        return { mode: 'breakdown', ...emptyResult([{ code: 'EMPTY_INPUT', message: 'The upload contains no files' }]) };
    }
    if (files.length === 1) {
        return ingest(files[0].buffer, { filename: files[0].filename, format: files[0].format });
    }
    for (const file of files) {
        if (isLegacyXls(file.buffer)) {
            return {
                mode: 'breakdown',
                ...emptyResult([
                    {
                        code: 'UNSUPPORTED_FORMAT',
                        message: `"${
                            file.filename ?? 'One of the files'
                        }" is a legacy .xls (BIFF) or password-protected workbook — re-save it as an unencrypted .xlsx or export .csv`,
                    },
                ]),
            };
        }
    }

    const sheets: SheetRows[] = [];
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const label = file.filename ?? `File ${i + 1}`;
        const format = detectFormat(file.buffer, file.filename, file.format);
        if (format === 'csv') {
            sheets.push({ name: label, rows: rowsFromCsv(file.buffer) });
        } else {
            for (const sheet of await sheetsFromXlsx(file.buffer)) {
                sheets.push({ name: `${label} › ${sheet.name}`, rows: sheet.rows });
            }
        }
    }
    return ingestSheets(sheets);
}
