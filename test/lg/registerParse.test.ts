/**
 * GOAL-3 R3 — cheque-register sheet parser (src/lg/registerParse.ts).
 *
 * Covers the file's data-quality landmines (GOAL-3 §4.3): the 1901-01-01 /
 * journal-0 "never paid" sentinels, ops PAID variants, dd/mm/yyyy TEXT ops
 * dates (column-scoped — never a general date fallback, GOAL.md §11.3),
 * ISO-numeric currency codes, trailing-space + mojibake padding, raw Excel
 * serial dates, and blank padding rows.
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { sheetsFromXlsx } from '../../src/lg/ingest';
import { parseDmyDate, parseRegisterSheet, REGISTER_REQUIRED_FIELDS } from '../../src/lg/registerParse';
import { RawRow } from '../../src/shared/models';

const FIXTURE = join(__dirname, '..', 'fixtures', 'lg', 'register-sample.xlsx');

/** Minimal register sheet: header + rows over just the columns the engine maps. */
const HEADERS: RawRow = [
    'etl_date',
    'c2_instr_type',
    'c6_chq_no',
    'c9_amount',
    'c10_payee_name',
    'c12_status',
    'Ops Remark',
    'Ops Journal',
    'Ops Date',
    'c13_issued_branch',
    'c16_issued_date',
    'c17_issued_post_dt',
    'c19_issued_jrnl_no',
    'c25_matchd_post_dt',
    'c27_matchd_jrnl_no',
    'c29_stpd_rsn_cd',
    'c31_cncld_date',
    'c37_purchaser_name',
    'c38_beneficiary_name',
    'c48_currency',
];

function row(overrides: Partial<Record<string, unknown>>): RawRow {
    const cells: RawRow = new Array(HEADERS.length).fill(null);
    // Baseline: a plausible paid cheque.
    const base: Record<string, unknown> = {
        etl_date: new Date('2026-02-03T00:00:00Z'),
        c2_instr_type: 'PO',
        c6_chq_no: 1001,
        c9_amount: 100,
        c12_status: '02',
        c13_issued_branch: '001',
        c16_issued_date: new Date('2025-03-10T00:00:00Z'),
        c17_issued_post_dt: new Date('2025-03-10T00:00:00Z'),
        c19_issued_jrnl_no: 5001,
        c25_matchd_post_dt: new Date('2025-04-01T00:00:00Z'),
        c27_matchd_jrnl_no: 6001,
        ...overrides,
    };
    for (const [key, value] of Object.entries(base)) {
        const idx = HEADERS.indexOf(key);
        if (idx >= 0) {
            cells[idx] = value as RawRow[number];
        }
    }
    return cells;
}

describe('parseDmyDate', () => {
    test('parses strict day-first d[d]/m[m]/yyyy text', () => {
        expect(parseDmyDate('15/08/2025')).toBe('2025-08-15');
        expect(parseDmyDate('01/01/2004')).toBe('2004-01-01');
        // Single-digit variants measured in the real file (Task 12 calibration).
        expect(parseDmyDate('8/1/2023')).toBe('2023-01-08');
        expect(parseDmyDate('11/1/2023')).toBe('2023-01-11');
        expect(parseDmyDate('2/02/2023')).toBe('2023-02-02');
    });

    test('rejects impossible calendar dates and other shapes', () => {
        expect(parseDmyDate('13/13/2025')).toBeUndefined();
        expect(parseDmyDate('31/02/2025')).toBeUndefined();
        expect(parseDmyDate('2025-08-15')).toBeUndefined(); // ISO is not this parser's job
        expect(parseDmyDate('8/15/2025')).toBeUndefined(); // month-first is a landmine, not a date
        expect(parseDmyDate('09//04/2024')).toBeUndefined(); // double-slash typo — surfaced, not guessed
        expect(parseDmyDate('100465347')).toBeUndefined(); // a journal number pasted in the date column
        expect(parseDmyDate('')).toBeUndefined();
    });
});

describe('parseRegisterSheet', () => {
    test('parses a register row into RegisterCheque', () => {
        const { cheques, errors } = parseRegisterSheet([HEADERS, row({})], 'Sheet1');
        expect(errors).toHaveLength(0);
        expect(cheques).toHaveLength(1);
        const c = cheques[0];
        expect(c.chequeNumber).toBe('1001');
        expect(c.amountFils).toBe(100000);
        expect(c.status).toBe('02');
        expect(c.issuedDate).toBe('2025-03-10');
        expect(c.issuedPostDate).toBe('2025-03-10');
        expect(c.issuedJournal).toBe('5001');
        expect(c.matchedPostDate).toBe('2025-04-01');
        expect(c.matchedJournal).toBe('6001');
        expect(c.issuedBranch).toBe('001');
        expect(c.instrument).toBe('PO');
        expect(c.opsPaid).toBe(false);
        expect(c.rowNumber).toBe(1);
        expect(c.sheet).toBe('Sheet1');
    });

    test('sentinel matched date/journal become undefined', () => {
        const { cheques } = parseRegisterSheet([
            HEADERS,
            row({ c25_matchd_post_dt: new Date('1901-01-01T00:00:00Z'), c27_matchd_jrnl_no: 0, c12_status: '01' }),
        ]);
        expect(cheques[0].matchedPostDate).toBeUndefined();
        expect(cheques[0].matchedJournal).toBeUndefined();
    });

    test('ops PAID variants normalise to opsPaid', () => {
        const variants = ['PAID', 'Paid ', 'paid', ' PAID  '];
        for (const remark of variants) {
            const { cheques } = parseRegisterSheet([HEADERS, row({ 'Ops Remark': remark })]);
            expect(cheques[0].opsPaid).toBe(true);
            expect(cheques[0].opsRemark).toBe(remark.trim());
        }
        const { cheques } = parseRegisterSheet([HEADERS, row({ 'Ops Remark': 'REVERSED' })]);
        expect(cheques[0].opsPaid).toBe(false);
    });

    test('ops date parses dd/mm/yyyy text; bad shapes become a row error, not a guess', () => {
        const ok = parseRegisterSheet([HEADERS, row({ 'Ops Date': '15/08/2025' })]);
        expect(ok.cheques[0].opsDate).toBe('2025-08-15');
        expect(ok.errors).toHaveLength(0);

        const bad = parseRegisterSheet([HEADERS, row({ 'Ops Date': '13/13/2025' })]);
        expect(bad.cheques[0].opsDate).toBeUndefined();
        expect(bad.errors).toEqual([expect.objectContaining({ code: 'BAD_DATE', row: 1 })]);
    });

    test('currency 48 maps to BHD; text fields trimmed and mojibake stripped', () => {
        const { cheques } = parseRegisterSheet([
            HEADERS,
            row({ c48_currency: '48', c10_payee_name: 'JOHN DOE  �', c38_beneficiary_name: ' SOME ONE � ' }),
        ]);
        expect(cheques[0].currency).toBe('BHD');
        expect(cheques[0].payee).toBe('JOHN DOE');
        expect(cheques[0].beneficiary).toBe('SOME ONE');
    });

    test('raw Excel serial in a date column coerces', () => {
        const { cheques } = parseRegisterSheet([HEADERS, row({ c16_issued_date: 45458 })]);
        expect(cheques[0].issuedDate).toBe('2024-06-15');
    });

    test('a bad amount is a row error and the row is dropped', () => {
        const { cheques, errors } = parseRegisterSheet([HEADERS, row({ c9_amount: 'n/a' })]);
        expect(cheques).toHaveLength(0);
        expect(errors).toEqual([expect.objectContaining({ code: 'BAD_AMOUNT', row: 1 })]);
    });

    test('blank padding rows are skipped silently', () => {
        const blank: RawRow = new Array(HEADERS.length).fill(null);
        blank[0] = ' ';
        const { cheques, errors } = parseRegisterSheet([HEADERS, row({}), blank, blank]);
        expect(cheques).toHaveLength(1);
        expect(errors).toHaveLength(0);
    });

    test('required register fields are exported for role detection', () => {
        expect(REGISTER_REQUIRED_FIELDS.length).toBeGreaterThanOrEqual(4);
    });

    test('parses the fixture register sheet: 11 cheques with the designed shapes', async () => {
        const sheets = await sheetsFromXlsx(readFileSync(FIXTURE));
        const register = sheets.find((s) => s.name === 'Sheet1')!;
        const { cheques, errors } = parseRegisterSheet(register.rows, 'Sheet1');
        expect(errors).toHaveLength(0);
        expect(cheques).toHaveLength(11);

        const byChq = new Map(cheques.map((c) => [c.chequeNumber, c]));
        expect(byChq.get('1002')!.matchedPostDate).toBeUndefined();
        expect(byChq.get('1002')!.currency).toBe('BHD'); // '48' mapped
        expect(byChq.get('1003')!.issuedDate).toBe('2024-06-15'); // raw serial coerced
        expect(byChq.get('1004')!.opsPaid).toBe(true);
        expect(byChq.get('1004')!.opsJournal).toBe('7001');
        expect(byChq.get('1004')!.opsDate).toBe('2025-08-15'); // dd/mm/yyyy text
        expect(byChq.get('1005')!.payee).toBe('JOHN DOE'); // trimmed + de-mojibaked
        expect(byChq.get('1011')!.amountFils).toBe(30000);
        expect(cheques.reduce((s, c) => s + c.amountFils, 0)).toBe(2700750); // Σ register amounts
    });
});

/**
 * Wrong-value tracking: non-empty cells that fail their column's format are
 * saved (raw value + column) on the error. Wrong dates in the register's
 * non-key columns no longer vanish silently — the cheque still parses and the
 * wrongness is tracked with rowParsed. Sentinels are markers, never "wrong".
 */
describe('wrong-value tracking', () => {
    test('a wrong amount saves the raw value and column (the cheque cannot parse)', () => {
        const { cheques, errors } = parseRegisterSheet([HEADERS, row({ c9_amount: 'oops' })], 'Sheet1');
        expect(cheques).toHaveLength(0);
        expect(errors).toContainEqual(
            expect.objectContaining({ code: 'BAD_AMOUNT', value: 'oops', column: 'D', columnHeader: 'c9_amount' })
        );
    });

    test('an ops-date typo is tracked with value+column and the cheque still parses', () => {
        const { cheques, errors } = parseRegisterSheet(
            [HEADERS, row({ 'Ops Remark': 'PAID', 'Ops Date': '09//04/2024' })],
            'Sheet1'
        );
        expect(cheques).toHaveLength(1);
        expect(cheques[0].opsDate).toBeUndefined();
        expect(errors).toContainEqual(
            expect.objectContaining({
                code: 'BAD_DATE',
                value: '09//04/2024',
                column: 'I',
                columnHeader: 'Ops Date',
                rowParsed: true,
            })
        );
    });

    test('garbage in the issued/matched/cancel date columns is tracked, not silently dropped', () => {
        const { cheques, errors } = parseRegisterSheet(
            [HEADERS, row({ c16_issued_date: 'garbage', c25_matchd_post_dt: '??', c31_cncld_date: 'x1' })],
            'Sheet1'
        );
        expect(cheques).toHaveLength(1);
        expect(cheques[0].issuedDate).toBeUndefined();
        expect(cheques[0].matchedPostDate).toBeUndefined();
        expect(cheques[0].cancelDate).toBeUndefined();
        expect(errors).toContainEqual(
            expect.objectContaining({ code: 'BAD_DATE', value: 'garbage', column: 'K', columnHeader: 'c16_issued_date', rowParsed: true })
        );
        expect(errors).toContainEqual(
            expect.objectContaining({ code: 'BAD_DATE', value: '??', column: 'N', columnHeader: 'c25_matchd_post_dt', rowParsed: true })
        );
        expect(errors).toContainEqual(
            expect.objectContaining({ code: 'BAD_DATE', value: 'x1', column: 'Q', columnHeader: 'c31_cncld_date', rowParsed: true })
        );
    });

    test('the never-paid sentinels are NOT wrong values', () => {
        const { cheques, errors } = parseRegisterSheet(
            [HEADERS, row({ c25_matchd_post_dt: new Date('1901-01-01T00:00:00Z'), c27_matchd_jrnl_no: 0 })],
            'Sheet1'
        );
        expect(cheques).toHaveLength(1);
        expect(cheques[0].matchedPostDate).toBeUndefined();
        expect(cheques[0].matchedJournal).toBeUndefined();
        expect(errors).toHaveLength(0);
    });

    test('empty optional date cells are missing, not wrong — no tracking entries', () => {
        const { cheques, errors } = parseRegisterSheet(
            [HEADERS, row({ c16_issued_date: null, c25_matchd_post_dt: '', c31_cncld_date: '   ' })],
            'Sheet1'
        );
        expect(cheques).toHaveLength(1);
        expect(errors).toHaveLength(0);
    });
});
