/**
 * LG reconciliation — parse & normalise raw breakdown rows into LedgerPostings
 * (GOAL.md §4 F2). Pure functions only: no I/O, no xlsx library. The ingest layer
 * (ingest.ts, F1) turns a file into `RawRow[]` and hands it here.
 */

import {
    AMOUNT_SCALE,
    CanonicalField,
    filsToBhd,
    LedgerPosting,
    ParseError,
    ParseResult,
    PostingDirection,
    RawCell,
    RawRow,
} from '../shared/models';

/**
 * Normalises a header cell for matching: lowercase, and strip whitespace, the
 * `_x000D_` carriage-return escape that shows up in the source amount headers
 * (e.g. "Amount\r\n(BHD)"), and underscores. Keeps parentheses and '?'.
 */
export function normalizeHeader(value: RawCell): string {
    return String(value ?? '')
        .toLowerCase()
        .replace(/_x000d_/g, '')
        .replace(/[\s_]+/g, '');
}

/** Accepted normalised header aliases for each canonical field. */
const HEADER_ALIASES: Record<CanonicalField, string[]> = {
    entity: ['entity'],
    branchNumber: ['branchnumber'],
    sbu: ['sbu'],
    level6: ['level6'],
    level3: ['level3'],
    level0: ['level0'],
    glDesc: ['gldesc'],
    glName: ['dynamicsglname', 'glname'],
    gl: ['gl'],
    accountNumber: ['accountnumber'],
    postDate: ['postdate'],
    postTime: ['posttime'],
    valueDate: ['valedate', 'valuedate'],
    source: ['source'],
    logDescription: ['logdescription'],
    currency: ['ccy', 'currency'],
    amountFcy: ['amount(fcy)'],
    amountLcy: ['amount(lcy)'],
    amountBhd: ['amount(bhd)'],
    journalNumber: ['journalnumber'],
    sequence: ['sequence'],
    userId: ['userid'],
    username: ['username'],
};

/** Fields whose header must be present for the file to be reconcilable. */
export const REQUIRED_FIELDS: CanonicalField[] = [
    'entity',
    'branchNumber',
    'gl',
    'postDate',
    'logDescription',
    'currency',
    'amountBhd',
    'journalNumber',
];

export type ColumnIndex = Partial<Record<CanonicalField, number>>;

export interface HeaderResult {
    columns: ColumnIndex;
    errors: ParseError[];
}

/** Maps a header row to column indexes, reporting any missing required columns. */
export function mapHeaders(headerRow: RawRow): HeaderResult {
    const normalized = headerRow.map(normalizeHeader);
    const columns: ColumnIndex = {};
    for (const field of Object.keys(HEADER_ALIASES) as CanonicalField[]) {
        const aliases = HEADER_ALIASES[field];
        const idx = normalized.findIndex((h) => aliases.includes(h));
        if (idx >= 0) {
            columns[field] = idx;
        }
    }
    const errors: ParseError[] = [];
    for (const field of REQUIRED_FIELDS) {
        if (columns[field] === undefined) {
            errors.push({ code: 'MISSING_HEADER', field, message: `Required column "${field}" was not found in the header row` });
        }
    }
    return { columns, errors };
}

/**
 * Parses a money value into signed integer fils (value × 1000, rounded).
 * Accepts numbers or strings; strips thousands separators and treats a
 * parenthesised value as negative. Returns undefined when it isn't a number.
 */
export function parseAmountToFils(value: RawCell): number | undefined {
    if (value === null || value === undefined || value === '') {
        return undefined;
    }
    if (typeof value === 'number') {
        return Number.isFinite(value) ? Math.round(value * AMOUNT_SCALE) : undefined;
    }
    let text = String(value).trim();
    if (text === '') {
        return undefined;
    }
    let sign = 1;
    if (/^\(.*\)$/.test(text)) {
        sign = -1;
        text = text.slice(1, -1);
    }
    text = text.replace(/,/g, '').trim();
    const n = Number(text);
    if (!Number.isFinite(n)) {
        return undefined;
    }
    return Math.round(n * AMOUNT_SCALE) * sign;
}

function pad2(n: number): string {
    return String(n).padStart(2, '0');
}

/**
 * Coerces a cell into an ISO date (yyyy-mm-dd). Handles Date objects (from the
 * xlsx reader), ISO/`yyyy-mm-dd` strings, and Excel serial-day numbers.
 */
export function coerceDate(value: RawCell): string | undefined {
    if (value === null || value === undefined || value === '') {
        return undefined;
    }
    if (value instanceof Date && !isNaN(value.getTime())) {
        return `${value.getUTCFullYear()}-${pad2(value.getUTCMonth() + 1)}-${pad2(value.getUTCDate())}`;
    }
    if (typeof value === 'number' && Number.isFinite(value)) {
        // Excel serial date: day 1 = 1900-01-01, with the well-known 1900 leap bug (offset 25569 to Unix epoch).
        const ms = Math.round((value - 25569) * 86400 * 1000);
        const d = new Date(ms);
        if (isNaN(d.getTime())) {
            return undefined;
        }
        return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
    }
    const text = String(value).trim();
    const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) {
        return `${iso[1]}-${iso[2]}-${iso[3]}`;
    }
    const parsed = new Date(text);
    if (!isNaN(parsed.getTime())) {
        return `${parsed.getUTCFullYear()}-${pad2(parsed.getUTCMonth() + 1)}-${pad2(parsed.getUTCDate())}`;
    }
    return undefined;
}

/** Extracts the leading 6-digit posting-type code from a log description. */
export function extractLogCode(logDescription: RawCell): string | undefined {
    const m = String(logDescription ?? '').trim().match(/^(\d{6})/);
    return m ? m[1] : undefined;
}

/** Posting-type codes that indicate direction, used to break ties on zero amounts. */
const DEBIT_LOG_CODES = new Set(['020050']);
const CREDIT_LOG_CODES = new Set(['020030']);

/** Direction from signed fils; falls back to the log code when the amount is zero. */
export function deriveDirection(amountBhdFils: number, logCode?: string): PostingDirection | undefined {
    if (amountBhdFils > 0) {
        return 'debit';
    }
    if (amountBhdFils < 0) {
        return 'credit';
    }
    if (logCode && DEBIT_LOG_CODES.has(logCode)) {
        return 'debit';
    }
    if (logCode && CREDIT_LOG_CODES.has(logCode)) {
        return 'credit';
    }
    return undefined;
}

function str(value: RawCell): string | undefined {
    if (value === null || value === undefined) {
        return undefined;
    }
    const s = String(value).trim();
    return s === '' ? undefined : s;
}

export interface RowResult {
    posting?: LedgerPosting;
    errors: ParseError[];
}

/** Normalises one data row into a LedgerPosting, collecting per-row errors. */
export function normalizeRow(row: RawRow, columns: ColumnIndex, rowNumber: number): RowResult {
    const errors: ParseError[] = [];
    const cell = (field: CanonicalField): RawCell => {
        const idx = columns[field];
        return idx === undefined ? undefined : row[idx];
    };

    const branchNumber = str(cell('branchNumber'));
    const gl = str(cell('gl'));
    const journalNumber = str(cell('journalNumber'));
    const currency = str(cell('currency'));
    const logDescription = str(cell('logDescription')) ?? '';

    const postDate = coerceDate(cell('postDate'));
    if (postDate === undefined) {
        errors.push({ code: 'BAD_DATE', field: 'postDate', row: rowNumber, message: 'Post Date is missing or not a valid date' });
    }

    const amountBhdFils = parseAmountToFils(cell('amountBhd'));
    if (amountBhdFils === undefined) {
        errors.push({ code: 'BAD_AMOUNT', field: 'amountBhd', row: rowNumber, message: 'Amount (BHD) is missing or not a number' });
    }

    for (const [field, value] of [
        ['branchNumber', branchNumber],
        ['gl', gl],
        ['journalNumber', journalNumber],
        ['currency', currency],
    ] as [CanonicalField, string | undefined][]) {
        if (value === undefined) {
            errors.push({ code: 'MISSING_FIELD', field, row: rowNumber, message: `Required value "${field}" is empty` });
        }
    }

    const logCode = extractLogCode(logDescription);
    const direction = amountBhdFils === undefined ? undefined : deriveDirection(amountBhdFils, logCode);
    if (amountBhdFils !== undefined && direction === undefined) {
        errors.push({ code: 'ZERO_AMOUNT', field: 'amountBhd', row: rowNumber, message: 'Amount is zero with no directional log code; cannot classify as debit or credit' });
    }

    if (errors.length > 0 || amountBhdFils === undefined || direction === undefined || !branchNumber || !gl || !journalNumber || !currency || !postDate) {
        return { errors };
    }

    const posting: LedgerPosting = {
        entity: str(cell('entity')) ?? '',
        branchNumber,
        sbu: str(cell('sbu')),
        level6: str(cell('level6')),
        level3: str(cell('level3')),
        level0: str(cell('level0')),
        glDesc: str(cell('glDesc')),
        glName: str(cell('glName')),
        gl,
        accountNumber: str(cell('accountNumber')),
        postDate,
        postTime: str(cell('postTime')),
        valueDate: coerceDate(cell('valueDate')),
        source: str(cell('source')),
        logDescription,
        logCode,
        currency,
        amountFcyFils: parseAmountToFils(cell('amountFcy')),
        amountLcyFils: parseAmountToFils(cell('amountLcy')),
        amountBhdFils,
        amountBhd: filsToBhd(amountBhdFils),
        direction,
        journalNumber,
        sequence: str(cell('sequence')),
        userId: str(cell('userId')),
        username: str(cell('username')),
        rowNumber,
    };
    return { posting, errors };
}

/** Whether a raw row is entirely empty (all cells blank). */
function isBlankRow(row: RawRow): boolean {
    return row.every((c) => c === null || c === undefined || String(c).trim() === '');
}

/**
 * Parses raw rows (row 0 = header) into a ParseResult. Header problems short-circuit
 * row parsing (there's nothing safe to map); row problems are collected, not thrown.
 */
export function parseRows(rows: RawRow[]): ParseResult {
    const empty: ParseResult = {
        postings: [],
        errors: [],
        summary: { dataRows: 0, parsed: 0, debitCount: 0, creditCount: 0, netFils: 0, currencies: [], branches: [] },
    };
    if (!rows || rows.length === 0) {
        return { ...empty, errors: [{ code: 'EMPTY_INPUT', message: 'The file contains no rows' }] };
    }

    const header = mapHeaders(rows[0]);
    if (header.errors.length > 0) {
        return { ...empty, errors: header.errors };
    }

    const postings: LedgerPosting[] = [];
    const errors: ParseError[] = [];
    let dataRows = 0;
    let debitCount = 0;
    let creditCount = 0;
    let netFils = 0;
    const currencies = new Set<string>();
    const branches = new Set<string>();

    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (isBlankRow(row)) {
            continue;
        }
        dataRows++;
        const { posting, errors: rowErrors } = normalizeRow(row, header.columns, dataRows);
        errors.push(...rowErrors);
        if (posting) {
            postings.push(posting);
            netFils += posting.amountBhdFils;
            if (posting.direction === 'debit') {
                debitCount++;
            } else {
                creditCount++;
            }
            currencies.add(posting.currency);
            branches.add(posting.branchNumber);
        }
    }

    return {
        postings,
        errors,
        summary: {
            dataRows,
            parsed: postings.length,
            debitCount,
            creditCount,
            netFils,
            currencies: [...currencies].sort(),
            branches: [...branches].sort(),
        },
    };
}
