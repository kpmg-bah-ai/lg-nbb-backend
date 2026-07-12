/**
 * GOAL-3 R2 — parse a Nostro/BGL ledger-statement worksheet (the `Credit` /
 * `Debit` sheets of the register-family workbook) into ParsedPostings. Pure.
 *
 * One schema serves both sheets: DIRECTION COMES PER ROW from which amount
 * column is populated (GOAL-3 §4.2) — `Transaction Credit Amount` populated ⇒
 * credit (stored negative, engine convention), `Transaction Debit Amount`
 * populated ⇒ debit (file holds debits negative ⇒ stored positive |value|).
 * Sheet names are never trusted for direction.
 *
 * Real-header quirks tolerated (Task 0): `Transaction Date` appears twice
 * (first occurrence wins), col 8 has a blank header, and the Debit sheet's
 * `reconciled` column sits after six blank headers.
 */

import { filsToBhd, ParsedPosting, ParseError, RawCell, RawRow } from '../shared/models';
import { cellTrack, coerceDate, excelColumn, extractLogCode, normalizeHeader, parseAmountToFils, rawCellText } from './parse';

type StatementField =
    | 'transactionDate'
    | 'postingDate'
    | 'gl'
    | 'journalNumber'
    | 'accountName'
    | 'description'
    | 'chequeNumber'
    | 'creditAmount'
    | 'debitAmount'
    | 'transactionType'
    | 'teller'
    | 'branch'
    | 'statedEod'
    | 'statedPrev'
    | 'detailedDescription'
    | 'sequence'
    | 'reconciled';

/** normalizeHeader('Nostro/BGL Account') — exported for role detection (detect.ts). */
export const STATEMENT_GL_ALIAS = 'nostro/bglaccount';

const STATEMENT_ALIASES: Record<StatementField, string[]> = {
    transactionDate: ['transactiondate'],
    postingDate: ['postingdate'],
    gl: [STATEMENT_GL_ALIAS],
    journalNumber: ['journalnumber'],
    accountName: ['accountname'],
    description: ['transactiondescription'],
    chequeNumber: ['chequenumber'],
    creditAmount: ['transactioncreditamount'],
    debitAmount: ['transactiondebitamount'],
    transactionType: ['transactiontype'],
    teller: ['teller'],
    branch: ['branch'],
    statedEod: ['enddateeodbalance'],
    statedPrev: ['previouseodbalance'],
    detailedDescription: ['detaileddescription'],
    sequence: ['sequencenumber'],
    reconciled: ['reconciled'],
};

export type StatementColumnIndex = Partial<Record<StatementField, number>>;

/**
 * Maps a header row to statement column indexes. A sheet qualifies (GOAL-3 §4.1)
 * when it carries transaction date, journal number, branch AND at least one of
 * the two amount columns. First occurrence wins for duplicated headers.
 */
export function mapStatementHeaders(headerRow: RawRow): StatementColumnIndex | undefined {
    const normalized = headerRow.map(normalizeHeader);
    const columns: StatementColumnIndex = {};
    for (const field of Object.keys(STATEMENT_ALIASES) as StatementField[]) {
        const idx = normalized.findIndex((h) => STATEMENT_ALIASES[field].includes(h));
        if (idx >= 0) {
            columns[field] = idx;
        }
    }
    const required =
        columns.transactionDate !== undefined &&
        columns.journalNumber !== undefined &&
        columns.branch !== undefined &&
        (columns.creditAmount !== undefined || columns.debitAmount !== undefined);
    return required ? columns : undefined;
}

/** How many leading rows to scan for the header (preamble tolerance, mirrors parse.ts). */
const HEADER_SCAN_ROWS = 10;

export function findStatementHeaderRow(rows: RawRow[], maxScan = HEADER_SCAN_ROWS): number {
    const limit = Math.min(rows.length, maxScan);
    for (let i = 0; i < limit; i++) {
        if (mapStatementHeaders(rows[i])) {
            return i;
        }
    }
    return -1;
}

function str(value: RawCell): string | undefined {
    if (value === null || value === undefined) {
        return undefined;
    }
    const s = String(value).trim();
    return s === '' ? undefined : s;
}

function isBlankRow(row: RawRow): boolean {
    return row.every((c) => c === null || c === undefined || String(c).trim() === '');
}

export interface StatementParseResult {
    postings: ParsedPosting[];
    errors: ParseError[];
    /** Non-blank data rows processed on this sheet (parsed + row-error rows). */
    dataRows: number;
}

/**
 * Parses one ledger-statement worksheet. `rowOffset` keeps rowNumber continuous
 * across the workbook's sheets (the caller parses Credit then Debit).
 */
export function parseStatementSheet(rows: RawRow[], sheetName?: string, rowOffset = 0): StatementParseResult {
    const errors: ParseError[] = [];
    const postings: ParsedPosting[] = [];
    const headerIndex = findStatementHeaderRow(rows);
    if (headerIndex < 0) {
        return {
            postings,
            errors: [
                {
                    code: 'MISSING_HEADER',
                    sheet: sheetName,
                    message:
                        'No ledger-statement header row found (need transaction date, journal number, branch and an amount column)',
                },
            ],
            dataRows: 0,
        };
    }
    const header = rows[headerIndex];
    const columns = mapStatementHeaders(header)!;
    const cell = (row: RawRow, field: StatementField): RawCell => {
        const idx = columns[field];
        return idx === undefined ? undefined : row[idx];
    };

    let rowNumber = rowOffset;
    for (let i = headerIndex + 1; i < rows.length; i++) {
        const row = rows[i];
        if (isBlankRow(row)) {
            continue; // blank padding rows (4,840 in the reference Debit sheet)
        }
        rowNumber++;
        const rowErrors: ParseError[] = [];

        const transactionDate = coerceDate(cell(row, 'transactionDate'));
        if (transactionDate === undefined) {
            rowErrors.push({
                code: 'BAD_DATE',
                field: 'postDate',
                row: rowNumber,
                sheet: sheetName,
                message: 'Transaction Date is missing or not a valid date',
                ...cellTrack(header, columns.transactionDate, cell(row, 'transactionDate')),
            });
        }

        const gl = str(cell(row, 'gl'));
        const branchNumber = str(cell(row, 'branch'));
        const journalNumber = str(cell(row, 'journalNumber'));
        for (const [field, column, value] of [
            ['gl', 'gl', gl],
            ['branchNumber', 'branch', branchNumber],
            ['journalNumber', 'journalNumber', journalNumber],
        ] as ['gl' | 'branchNumber' | 'journalNumber', StatementField, string | undefined][]) {
            if (value === undefined) {
                rowErrors.push({
                    code: 'MISSING_FIELD',
                    field,
                    row: rowNumber,
                    sheet: sheetName,
                    message: `Required value "${field}" is empty`,
                    ...cellTrack(header, columns[column], cell(row, column)),
                });
            }
        }

        // Direction per row: exactly one non-zero amount column (GOAL-3 §4.2).
        const creditRaw = cell(row, 'creditAmount');
        const debitRaw = cell(row, 'debitAmount');
        const creditFils = parseAmountToFils(creditRaw);
        const debitFils = parseAmountToFils(debitRaw);
        const hasCredit = creditFils !== undefined && creditFils !== 0;
        const hasDebit = debitFils !== undefined && debitFils !== 0;
        let amountBhdFils: number | undefined;
        let direction: 'debit' | 'credit' | undefined;
        if (hasCredit && hasDebit) {
            const cols = [columns.creditAmount, columns.debitAmount].filter((c): c is number => c !== undefined);
            rowErrors.push({
                code: 'AMBIGUOUS_DIRECTION',
                field: 'amountBhd',
                row: rowNumber,
                sheet: sheetName,
                message: 'Both the credit and debit amount columns are populated — cannot classify the row',
                value: `${rawCellText(creditRaw)} / ${rawCellText(debitRaw)}`,
                column: cols.map(excelColumn).join('/'),
                columnHeader: cols.map((c) => rawCellText(header[c]) ?? '').join(' / '),
            });
        } else if (hasCredit) {
            direction = 'credit';
            amountBhdFils = -Math.abs(creditFils);
        } else if (hasDebit) {
            direction = 'debit';
            amountBhdFils = Math.abs(debitFils); // debits are held negative in the file
        } else {
            // Point the error at whichever cell actually holds content (wrong
            // text or a zero); when both are empty the amounts are missing,
            // not wrong, and there is no cell to track.
            const offending = rawCellText(creditRaw) !== undefined ? ('creditAmount' as const) : rawCellText(debitRaw) !== undefined ? ('debitAmount' as const) : undefined;
            rowErrors.push({
                code: 'BAD_AMOUNT',
                field: 'amountBhd',
                row: rowNumber,
                sheet: sheetName,
                message: 'Neither amount column carries a non-zero value',
                ...(offending ? cellTrack(header, columns[offending], cell(row, offending)) : {}),
            });
        }

        // Wrong values in the non-key columns: the row still parses; the wrong
        // cell is tracked with its raw value + column. Empty cells are missing,
        // not wrong — no tracking entry.
        const tracked: ParseError[] = [];
        const trackWrong = (field: StatementField, code: 'BAD_DATE' | 'BAD_AMOUNT', what: string): void => {
            const raw = cell(row, field);
            if (rawCellText(raw) !== undefined) {
                tracked.push({
                    code,
                    field,
                    row: rowNumber,
                    sheet: sheetName,
                    message: `${what} — row kept, wrong value tracked`,
                    ...cellTrack(header, columns[field], raw),
                });
            }
        };
        const postingDate = coerceDate(cell(row, 'postingDate'));
        if (postingDate === undefined) {
            trackWrong('postingDate', 'BAD_DATE', 'Posting Date is not a valid date (row kept on the transaction date)');
        }
        const statedEodFils = parseAmountToFils(cell(row, 'statedEod'));
        if (statedEodFils === undefined) {
            trackWrong('statedEod', 'BAD_AMOUNT', 'End Date EoD Balance is not a number');
        }
        const statedPrevEodFils = parseAmountToFils(cell(row, 'statedPrev'));
        if (statedPrevEodFils === undefined) {
            trackWrong('statedPrev', 'BAD_AMOUNT', 'Previous EoD Balance is not a number');
        }

        if (rowErrors.length > 0 || !gl || !branchNumber || !journalNumber || !transactionDate || !direction) {
            // Row excluded by a critical error: keep the tracking, without rowParsed.
            errors.push(...rowErrors, ...tracked);
            continue;
        }
        for (const t of tracked) {
            t.rowParsed = true;
        }
        errors.push(...tracked);

        const description = str(cell(row, 'description')) ?? '';
        postings.push({
            entity: '', // no entity column in this layout (GOAL-3 §2.4 default)
            branchNumber,
            gl,
            postDate: postingDate ?? transactionDate,
            logDescription: description,
            logCode: extractLogCode(description),
            currency: 'BHD', // no currency column — default (GOAL-3 §2.5)
            amountBhdFils: amountBhdFils!,
            amountBhd: filsToBhd(amountBhdFils!),
            direction,
            journalNumber,
            sequence: str(cell(row, 'sequence')),
            rowNumber,
            sheet: sheetName,
            // Register-family extensions (GOAL-3 §4.2):
            transactionDate,
            chequeNumber: str(cell(row, 'chequeNumber')),
            transactionType: str(cell(row, 'transactionType')),
            teller: str(cell(row, 'teller')),
            accountName: str(cell(row, 'accountName')),
            detailedDescription: str(cell(row, 'detailedDescription')),
            statedEodFils,
            statedPrevEodFils,
            reconciledNote: str(cell(row, 'reconciled')),
        });
    }

    return { postings, errors, dataRows: rowNumber - rowOffset };
}
