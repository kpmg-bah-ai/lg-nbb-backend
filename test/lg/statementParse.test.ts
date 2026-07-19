/**
 * GOAL-3 R2 — ledger-statement sheet parser (src/lg/statementParse.ts).
 *
 * One schema for both Credit and Debit sheets: direction comes PER ROW from
 * which amount column is populated (debits held negative in their own column),
 * never from the sheet name. Captures the stated End Date EoD Balance and the
 * Detailed Description (batch Ref.# lists) for later stages.
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { sheetsFromXlsx } from '../../src/lg/ingest';
import { parseStatementSheet, STATEMENT_GL_ALIAS } from '../../src/lg/statementParse';
import { RawRow } from '../../src/shared/models';

const FIXTURE = join(__dirname, '..', 'fixtures', 'lg', 'register-sample.xlsx');

// The real Task-0 header shape, including the duplicate Transaction Date (cols
// 0 and 20) and the blank header at col 8.
const HEADERS: RawRow = [
    'Transaction Date',
    'Posting Date',
    'Nostro/BGL Account',
    'Journal Number',
    'New Ref',
    'MGR CHQ Number',
    'Account Name',
    'Transaction Description',
    '',
    'Cheque Number',
    'Transaction Credit Amount',
    'Transaction Debit Amount',
    'Transaction Type',
    'Teller',
    'Branch',
    'End Date EoD Balance',
    'Previous EoD Balance',
    'Detailed Description',
    'Outlet',
    'Authorization',
    'Transaction Date',
    'Transaction Time',
    'RRN',
    'Card Number',
    'Balance',
    'Sequence Number',
];

function row(overrides: Partial<Record<string, unknown>> = {}): RawRow {
    const base: Record<string, unknown> = {
        'Transaction Date': new Date('2025-03-10T00:00:00Z'),
        'Posting Date': new Date('2025-03-11T00:00:00Z'),
        'Nostro/BGL Account': '99801000',
        'Journal Number': 5001,
        'Account Name': 'SAMPLE CUSTOMER',
        'Transaction Description': 'DD ISSUED',
        'Transaction Credit Amount': 100,
        'Transaction Type': 'DD FROM DEP A/C',
        Teller: 'T123',
        Branch: '001',
        'End Date EoD Balance': 2730,
        'Previous EoD Balance': 2600,
        'Detailed Description': '',
        'Sequence Number': 1,
        ...overrides,
    };
    const cells: RawRow = new Array(HEADERS.length).fill(null);
    for (const [key, value] of Object.entries(base)) {
        const idx = HEADERS.indexOf(key); // first occurrence wins — mirrors the parser
        if (idx >= 0) {
            cells[idx] = value as RawRow[number];
        }
    }
    return cells;
}

describe('parseStatementSheet', () => {
    test('a populated credit column makes a credit posting (negative signed fils)', () => {
        const { postings, errors } = parseStatementSheet([HEADERS, row()], 'Credit');
        expect(errors).toHaveLength(0);
        expect(postings).toHaveLength(1);
        const p = postings[0];
        expect(p.direction).toBe('credit');
        expect(p.amountBhdFils).toBe(-100000);
        expect(p.transactionDate).toBe('2025-03-10'); // the key-leg date
        expect(p.postDate).toBe('2025-03-11');
        expect(p.gl).toBe('99801000');
        expect(p.journalNumber).toBe('5001');
        expect(p.branchNumber).toBe('001');
        expect(p.entity).toBe(''); // no entity column in this layout (§2.4 default)
        expect(p.currency).toBe('BHD'); // no currency column — default (§2.5)
        expect(p.transactionType).toBe('DD FROM DEP A/C');
        expect(p.accountName).toBe('SAMPLE CUSTOMER');
        expect(p.statedEodFils).toBe(2730000);
        expect(p.statedPrevEodFils).toBe(2600000);
        expect(p.sheet).toBe('Credit');
        expect(p.rowNumber).toBe(1);
    });

    test('a populated debit column (file-negative) makes a positive debit posting', () => {
        const { postings } = parseStatementSheet(
            [HEADERS, row({ 'Transaction Credit Amount': null, 'Transaction Debit Amount': -250.5 })],
            'Debit'
        );
        expect(postings[0].direction).toBe('debit');
        expect(postings[0].amountBhdFils).toBe(250500);
    });

    test('both amount columns non-zero is AMBIGUOUS_DIRECTION; both empty is BAD_AMOUNT', () => {
        const both = parseStatementSheet(
            [HEADERS, row({ 'Transaction Credit Amount': 100, 'Transaction Debit Amount': -50 })],
            'Credit'
        );
        expect(both.postings).toHaveLength(0);
        expect(both.errors).toEqual([expect.objectContaining({ code: 'AMBIGUOUS_DIRECTION', row: 1 })]);

        const neither = parseStatementSheet(
            [HEADERS, row({ 'Transaction Credit Amount': null, 'Transaction Debit Amount': null })],
            'Credit'
        );
        expect(neither.postings).toHaveLength(0);
        expect(neither.errors).toEqual([expect.objectContaining({ code: 'BAD_AMOUNT', row: 1 })]);
    });

    test('missing posting date falls back to the transaction date', () => {
        const { postings } = parseStatementSheet([HEADERS, row({ 'Posting Date': null })], 'Credit');
        expect(postings[0].postDate).toBe('2025-03-10');
    });

    test('required fields missing become row errors', () => {
        const noJournal = parseStatementSheet([HEADERS, row({ 'Journal Number': null })], 'Credit');
        expect(noJournal.postings).toHaveLength(0);
        expect(noJournal.errors).toEqual([
            expect.objectContaining({ code: 'MISSING_FIELD', field: 'journalNumber', row: 1 }),
        ]);

        const noDate = parseStatementSheet([HEADERS, row({ 'Transaction Date': null })], 'Credit');
        expect(noDate.postings).toHaveLength(0);
        expect(noDate.errors).toEqual([expect.objectContaining({ code: 'BAD_DATE', row: 1 })]);
    });

    test('detailed description and cheque number are carried for later stages', () => {
        const { postings } = parseStatementSheet(
            [
                HEADERS,
                row({
                    'Transaction Credit Amount': null,
                    'Transaction Debit Amount': -200,
                    'Detailed Description': 'DEBIT POSTING-20-Ref.# 5007,Ref.# 5008',
                    'Cheque Number': 1007,
                }),
            ],
            'Debit'
        );
        expect(postings[0].detailedDescription).toBe('DEBIT POSTING-20-Ref.# 5007,Ref.# 5008');
        expect(postings[0].chequeNumber).toBe('1007');
    });

    test('the reconciled disposition column is carried when present', () => {
        const headers: RawRow = [...HEADERS, '', '', '', '', '', '', 'reconciled'];
        const cells = row({ 'Transaction Credit Amount': null, 'Transaction Debit Amount': -10 });
        const withNote: RawRow = [...cells, null, null, null, null, null, null, 'Datafix entry - not MC'];
        const { postings } = parseStatementSheet([headers, withNote], 'Debit');
        expect(postings[0].reconciledNote).toBe('Datafix entry - not MC');
    });

    test('rowOffset keeps row numbering continuous across sheets', () => {
        const { postings } = parseStatementSheet([HEADERS, row()], 'Debit', 12);
        expect(postings[0].rowNumber).toBe(13);
    });

    test('fixture Credit sheet: 12 postings summing to −3,675,500 fils', async () => {
        const sheets = await sheetsFromXlsx(readFileSync(FIXTURE));
        const credit = sheets.find((s) => s.name === 'Credit')!;
        const { postings, errors } = parseStatementSheet(credit.rows, 'Credit');
        expect(errors).toHaveLength(0);
        expect(postings).toHaveLength(12);
        expect(postings.every((p) => p.direction === 'credit')).toBe(true);
        expect(postings.reduce((s, p) => s + p.amountBhdFils, 0)).toBe(-3675500);
    });

    test('fixture Debit sheet: 4 postings summing to +910,000 fils, padding skipped', async () => {
        const sheets = await sheetsFromXlsx(readFileSync(FIXTURE));
        const debit = sheets.find((s) => s.name === 'Debit')!;
        const { postings, errors } = parseStatementSheet(debit.rows, 'Debit');
        expect(errors).toHaveLength(0);
        expect(postings).toHaveLength(4);
        expect(postings.every((p) => p.direction === 'debit')).toBe(true);
        expect(postings.reduce((s, p) => s + p.amountBhdFils, 0)).toBe(910000);
        const batch = postings.find((p) => p.journalNumber === '8001')!;
        expect(batch.detailedDescription).toContain('Ref.# 5007');
    });

    test('exports the GL header alias for role detection', () => {
        expect(STATEMENT_GL_ALIAS).toBe('nostro/bglaccount');
    });
});

/**
 * Wrong-value tracking: non-empty cells that fail their column's format are
 * saved (raw value + column) on the error. Wrong values in the non-key columns
 * (posting date, stated balances) no longer degrade silently — the row still
 * parses and the wrongness is tracked with rowParsed.
 */
describe('wrong-value tracking', () => {
    test('a wrong Transaction Date saves value+column (row cannot join matching)', () => {
        const { postings, errors } = parseStatementSheet([HEADERS, row({ 'Transaction Date': 'JNL 123' })], 'Credit');
        expect(postings).toHaveLength(0);
        expect(errors).toContainEqual(
            expect.objectContaining({ code: 'BAD_DATE', value: 'JNL 123', column: 'A', columnHeader: 'Transaction Date' })
        );
    });

    test('a wrong Posting Date is tracked and the row still parses on the transaction date', () => {
        const { postings, errors } = parseStatementSheet([HEADERS, row({ 'Posting Date': '31/02/2025x' })], 'Credit');
        expect(postings).toHaveLength(1);
        expect(postings[0].postDate).toBe('2025-03-10'); // fallback to the transaction date
        expect(errors).toContainEqual(
            expect.objectContaining({
                code: 'BAD_DATE',
                value: '31/02/2025x',
                column: 'B',
                columnHeader: 'Posting Date',
                rowParsed: true,
            })
        );
    });

    test('wrong stated-balance cells are tracked and the row still parses', () => {
        const { postings, errors } = parseStatementSheet(
            [HEADERS, row({ 'End Date EoD Balance': 'abc', 'Previous EoD Balance': 'de' })],
            'Credit'
        );
        expect(postings).toHaveLength(1);
        expect(postings[0].statedEodFils).toBeUndefined();
        expect(postings[0].statedPrevEodFils).toBeUndefined();
        expect(errors).toContainEqual(
            expect.objectContaining({ code: 'BAD_AMOUNT', value: 'abc', column: 'P', rowParsed: true })
        );
        expect(errors).toContainEqual(
            expect.objectContaining({ code: 'BAD_AMOUNT', value: 'de', column: 'Q', rowParsed: true })
        );
    });

    test('non-numeric text in an amount column is the saved wrong value on BAD_AMOUNT', () => {
        const { postings, errors } = parseStatementSheet([HEADERS, row({ 'Transaction Credit Amount': 'n/a' })], 'Credit');
        expect(postings).toHaveLength(0);
        expect(errors).toContainEqual(
            expect.objectContaining({ code: 'BAD_AMOUNT', value: 'n/a', column: 'K', columnHeader: 'Transaction Credit Amount' })
        );
    });

    test('AMBIGUOUS_DIRECTION saves both populated amounts and both columns', () => {
        const { errors } = parseStatementSheet(
            [HEADERS, row({ 'Transaction Credit Amount': 100, 'Transaction Debit Amount': -100 })],
            'Credit'
        );
        expect(errors).toContainEqual(
            expect.objectContaining({ code: 'AMBIGUOUS_DIRECTION', value: '100 / -100', column: 'K/L' })
        );
    });

    test('empty stated-balance cells are missing, not wrong — no tracking entries', () => {
        const { postings, errors } = parseStatementSheet(
            [HEADERS, row({ 'End Date EoD Balance': null, 'Previous EoD Balance': '' })],
            'Credit'
        );
        expect(postings).toHaveLength(1);
        expect(errors).toHaveLength(0);
    });
});

describe('GOAL-8: zero-amount rows (VAT statement GLs)', () => {
    const VAT_HEADER: RawRow = [
        'Transaction Date', 'Posting Date', 'Nostro/BGL Account', 'Journal Number', 'Account Name',
        'Transaction Description', 'Cheque Number', 'Transaction Credit Amount', 'Transaction Debit Amount',
        'Transaction Type', 'Teller', 'Branch', 'End Date EoD Balance',
    ];
    const vatRow = (credit: unknown, debit: unknown): RawRow => [
        new Date('2023-01-09T00:00:00Z'), new Date('2023-01-09T00:00:00Z'), '8828010400010000', 'J1',
        'INPUT VAT RECEIVABLE MUBASHER - BHD', 'NPB MISC DEP DR', '', credit as never, debit as never,
        '01-Financial', 'System', '00001-Main Branch', 0.499,
    ];

    it('classifies a literal-zero amount row as ZERO_AMOUNT, not BAD_AMOUNT', () => {
        const { postings, errors } = parseStatementSheet([VAT_HEADER, vatRow(0, null)], 'VAT');
        expect(postings).toHaveLength(0); // zero rows don't move the running balance
        expect(errors).toEqual([expect.objectContaining({ code: 'ZERO_AMOUNT', row: 1 })]);
    });

    it('still reports non-numeric junk in an amount column as BAD_AMOUNT', () => {
        const { errors } = parseStatementSheet([VAT_HEADER, vatRow('abc', null)], 'VAT');
        expect(errors).toEqual([expect.objectContaining({ code: 'BAD_AMOUNT', row: 1 })]);
    });
});
