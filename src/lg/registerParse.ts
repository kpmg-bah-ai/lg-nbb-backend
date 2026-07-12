/**
 * GOAL-3 R3 — parse the cheque-register worksheet (`Sheet1`, the c0…c60 ETL
 * extract + ops working columns) into RegisterCheque rows. Pure functions.
 *
 * Data-quality rules measured from the reference file (GOAL-3 §4.3):
 *  - "never paid" sentinels: matched post date 1901-01-01 and journal 0 ⇒ undefined;
 *  - `Ops Date` arrives as dd/mm/yyyy TEXT — parsed ONLY by the column-scoped
 *    parseDmyDate (the global date coercion stays ISO-strict, GOAL.md §11.3);
 *  - ISO-numeric currency text ('48' ⇒ BHD);
 *  - text fields carry trailing spaces and U+FFFD mojibake padding — cleaned;
 *  - dates can arrive as Date objects, ISO strings or raw Excel serials;
 *  - blank padding rows (9,610 in the reference) are skipped silently.
 */

import { ParseError, RawCell, RawRow, RegisterCheque } from '../shared/models';
import { coerceDate, normalizeHeader, parseAmountToFils } from './parse';

type RegisterField =
    | 'etlDate'
    | 'instrument'
    | 'chequeNumber'
    | 'amount'
    | 'payee'
    | 'status'
    | 'opsRemark'
    | 'opsJournal'
    | 'opsDate'
    | 'issuedBranch'
    | 'issuedDate'
    | 'issuedPostDate'
    | 'issuedJournal'
    | 'matchedPostDate'
    | 'matchedJournal'
    | 'stopReason'
    | 'cancelDate'
    | 'purchaser'
    | 'beneficiary'
    | 'currency';

/** Normalised header aliases (real strings frozen by Task 0 in register-headers.json). */
const REGISTER_ALIASES: Record<RegisterField, string[]> = {
    etlDate: ['etldate'],
    instrument: ['c2instrtype'],
    chequeNumber: ['c6chqno'],
    amount: ['c9amount'],
    payee: ['c10payeename'],
    status: ['c12status'],
    opsRemark: ['opsremark'],
    opsJournal: ['opsjournal'],
    opsDate: ['opsdate'],
    issuedBranch: ['c13issuedbranch'],
    issuedDate: ['c16issueddate'],
    issuedPostDate: ['c17issuedpostdt'],
    issuedJournal: ['c19issuedjrnlno'],
    matchedPostDate: ['c25matchdpostdt'],
    matchedJournal: ['c27matchdjrnlno'],
    stopReason: ['c29stpdrsncd'],
    cancelDate: ['c31cnclddate'],
    purchaser: ['c37purchasername'],
    beneficiary: ['c38beneficiaryname'],
    currency: ['c48currency'],
};

/** Headers that must be present for a sheet to count as a cheque register (GOAL-3 §4.1). */
export const REGISTER_REQUIRED_FIELDS: RegisterField[] = [
    'chequeNumber',
    'amount',
    'issuedJournal',
    'matchedJournal',
];

/** "Never paid" sentinels used by the cheque system (file §3). */
export const SENTINEL_MATCHED_DATE = '1901-01-01';
export const SENTINEL_MATCHED_JOURNAL = '0';

/** ISO-numeric currency codes seen in older register rows (§11.5 — extend as confirmed). */
const CURRENCY_CODE_MAP: Record<string, string> = { '48': 'BHD' };

export type RegisterColumnIndex = Partial<Record<RegisterField, number>>;

/** Maps a header row to register column indexes; empty when required columns are missing. */
export function mapRegisterHeaders(headerRow: RawRow): RegisterColumnIndex | undefined {
    const normalized = headerRow.map(normalizeHeader);
    const columns: RegisterColumnIndex = {};
    for (const field of Object.keys(REGISTER_ALIASES) as RegisterField[]) {
        const idx = normalized.findIndex((h) => REGISTER_ALIASES[field].includes(h));
        if (idx >= 0) {
            columns[field] = idx;
        }
    }
    return REGISTER_REQUIRED_FIELDS.every((f) => columns[f] !== undefined) ? columns : undefined;
}

/** How many leading rows to scan for the header (preamble tolerance, mirrors parse.ts). */
const HEADER_SCAN_ROWS = 10;

export function findRegisterHeaderRow(rows: RawRow[], maxScan = HEADER_SCAN_ROWS): number {
    const limit = Math.min(rows.length, maxScan);
    for (let i = 0; i < limit; i++) {
        if (mapRegisterHeaders(rows[i])) {
            return i;
        }
    }
    return -1;
}

/**
 * Strict dd/mm/yyyy text date, calendar-checked — used ONLY for the register's
 * `Ops Date` column. Anything else (ISO, month-first, garbage) is undefined:
 * surface, don't guess (GOAL.md §11.3).
 */
export function parseDmyDate(text: string): string | undefined {
    const m = text.trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (!m) {
        return undefined;
    }
    const d = Number(m[1]);
    const mo = Number(m[2]);
    const y = Number(m[3]);
    const date = new Date(Date.UTC(y, mo - 1, d));
    if (date.getUTCFullYear() !== y || date.getUTCMonth() !== mo - 1 || date.getUTCDate() !== d) {
        return undefined;
    }
    return `${m[3]}-${m[2]}-${m[1]}`;
}

/** Trims and strips U+FFFD mojibake padding; empty ⇒ undefined. */
export function cleanText(value: RawCell): string | undefined {
    if (value === null || value === undefined) {
        return undefined;
    }
    const s = String(value).replace(/�/g, '').trim();
    return s === '' ? undefined : s;
}

/** Plain trimmed string; empty ⇒ undefined. */
function str(value: RawCell): string | undefined {
    if (value === null || value === undefined) {
        return undefined;
    }
    const s = String(value).trim();
    return s === '' ? undefined : s;
}

/** Status codes are zero-padded text ('01'…'05'); a numeric cell loses the pad — restore it. */
function statusText(value: RawCell): string | undefined {
    const s = str(value);
    return s !== undefined && /^\d$/.test(s) ? s.padStart(2, '0') : s;
}

function isBlankRow(row: RawRow): boolean {
    return row.every((c) => c === null || c === undefined || String(c).trim() === '');
}

export interface RegisterParseResult {
    cheques: RegisterCheque[];
    errors: ParseError[];
    /** Non-blank data rows processed (parsed + row-error rows). */
    dataRows: number;
}

/** Parses one register worksheet's raw rows (header + data) into RegisterCheques. */
export function parseRegisterSheet(rows: RawRow[], sheetName?: string): RegisterParseResult {
    const errors: ParseError[] = [];
    const cheques: RegisterCheque[] = [];
    const headerIndex = findRegisterHeaderRow(rows);
    if (headerIndex < 0) {
        return {
            cheques,
            errors: [
                {
                    code: 'MISSING_HEADER',
                    sheet: sheetName,
                    message: 'No cheque-register header row found (need cheque no, amount, issued/matched journals)',
                },
            ],
            dataRows: 0,
        };
    }
    const columns = mapRegisterHeaders(rows[headerIndex])!;
    const cell = (row: RawRow, field: RegisterField): RawCell => {
        const idx = columns[field];
        return idx === undefined ? undefined : row[idx];
    };

    let rowNumber = 0;
    for (let i = headerIndex + 1; i < rows.length; i++) {
        const row = rows[i];
        if (isBlankRow(row)) {
            continue; // blank padding rows — skipped silently (GOAL-3 §4.3)
        }
        rowNumber++;

        const amountFils = parseAmountToFils(cell(row, 'amount'));
        if (amountFils === undefined) {
            errors.push({
                code: 'BAD_AMOUNT',
                row: rowNumber,
                sheet: sheetName,
                message: 'Register cheque amount (c9_amount) is missing or not a number',
            });
            continue;
        }

        // Payment key leg with "never paid" sentinels collapsed to undefined.
        let matchedPostDate = coerceDate(cell(row, 'matchedPostDate'));
        let matchedJournal = str(cell(row, 'matchedJournal'));
        if (matchedPostDate === SENTINEL_MATCHED_DATE) {
            matchedPostDate = undefined;
        }
        if (matchedJournal === SENTINEL_MATCHED_JOURNAL) {
            matchedJournal = undefined;
        }

        // Ops Date: dd/mm/yyyy TEXT in the file; tolerate Date/serial/ISO cells too.
        const opsDateRaw = cell(row, 'opsDate');
        let opsDate: string | undefined;
        if (opsDateRaw !== null && opsDateRaw !== undefined && String(opsDateRaw).trim() !== '') {
            opsDate = typeof opsDateRaw === 'string' ? parseDmyDate(opsDateRaw) : coerceDate(opsDateRaw);
            if (opsDate === undefined) {
                errors.push({
                    code: 'BAD_DATE',
                    row: rowNumber,
                    sheet: sheetName,
                    message: `Ops Date "${String(opsDateRaw).trim()}" is not a valid dd/mm/yyyy date`,
                });
            }
        }

        const opsRemark = cleanText(cell(row, 'opsRemark'));
        const currencyRaw = cleanText(cell(row, 'currency'));

        cheques.push({
            instrument: str(cell(row, 'instrument')),
            chequeNumber: str(cell(row, 'chequeNumber')),
            amountFils,
            payee: cleanText(cell(row, 'payee')),
            status: statusText(cell(row, 'status')),
            issuedDate: coerceDate(cell(row, 'issuedDate')),
            issuedPostDate: coerceDate(cell(row, 'issuedPostDate')),
            issuedBranch: str(cell(row, 'issuedBranch')),
            issuedJournal: str(cell(row, 'issuedJournal')),
            matchedPostDate,
            matchedJournal,
            stopReason: str(cell(row, 'stopReason')),
            cancelDate: coerceDate(cell(row, 'cancelDate')),
            purchaser: cleanText(cell(row, 'purchaser')),
            beneficiary: cleanText(cell(row, 'beneficiary')),
            currency: currencyRaw ? CURRENCY_CODE_MAP[currencyRaw] ?? currencyRaw : undefined,
            opsRemark,
            opsPaid: opsRemark !== undefined && opsRemark.toUpperCase() === 'PAID',
            opsJournal: str(cell(row, 'opsJournal')),
            opsDate,
            rowNumber,
            sheet: sheetName,
        });
    }

    return { cheques, errors, dataRows: rowNumber };
}
