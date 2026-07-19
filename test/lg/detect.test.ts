/**
 * GOAL-3 R1 — sheet-role detection and run-mode routing (src/lg/detect.ts +
 * ingest.ts). A workbook either belongs to the breakdown family (existing
 * 24-col schema) or the register family (ledger-statement sheets + cheque
 * register); mixing families or shipping half a register input is rejected
 * loudly, never guessed (GOAL-3 §4.1).
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import * as ExcelJS from 'exceljs';
import { detectSheetRole, resolveMode } from '../../src/lg/detect';
import { ingest } from '../../src/lg/ingest';
import { RawRow, SheetRole } from '../../src/shared/models';

const FIXTURE = join(__dirname, '..', 'fixtures', 'lg', 'register-sample.xlsx');

const BREAKDOWN_HEADER: RawRow = [
    'entity', 'Branch Number', 'gl', 'Post Date', 'Log description', 'ccy', 'Amount (BHD)', 'Journal Number',
];
const STATEMENT_HEADER: RawRow = [
    'Transaction Date', 'Posting Date', 'Nostro/BGL Account', 'Journal Number', 'Branch',
    'Transaction Credit Amount', 'Transaction Debit Amount', 'End Date EoD Balance',
];
const REGISTER_HEADER: RawRow = [
    'etl_date', 'c6_chq_no', 'c9_amount', 'c12_status', 'c17_issued_post_dt', 'c19_issued_jrnl_no',
    'c25_matchd_post_dt', 'c27_matchd_jrnl_no', 'Ops Remark',
];

async function workbookBuffer(sheets: { name: string; rows: unknown[][] }[]): Promise<Buffer> {
    const workbook = new ExcelJS.Workbook();
    for (const sheet of sheets) {
        const ws = workbook.addWorksheet(sheet.name);
        ws.addRows(sheet.rows);
    }
    return Buffer.from(await workbook.xlsx.writeBuffer());
}

describe('detectSheetRole', () => {
    it('recognises all four roles', () => {
        expect(detectSheetRole([BREAKDOWN_HEADER])).toBe('breakdown');
        expect(detectSheetRole([STATEMENT_HEADER])).toBe('ledgerStatement');
        expect(detectSheetRole([REGISTER_HEADER])).toBe('register');
        expect(detectSheetRole([['item', 'comment'], ['x', 'y']])).toBe('unknown');
    });

    it('tolerates preamble rows before the header', () => {
        expect(detectSheetRole([['Some title'], [], REGISTER_HEADER])).toBe('register');
    });
});

describe('resolveMode', () => {
    const modeOf = (roles: SheetRole[]) => resolveMode(roles);

    it('register + ledger statement(s) resolve to register mode', () => {
        expect(modeOf(['ledgerStatement', 'register', 'ledgerStatement']).mode).toBe('register');
        expect(modeOf(['ledgerStatement', 'register', 'unknown']).mode).toBe('register');
    });

    it('breakdown sheets resolve to breakdown mode', () => {
        expect(modeOf(['breakdown']).mode).toBe('breakdown');
        expect(modeOf(['breakdown', 'unknown']).mode).toBe('breakdown');
    });

    it('mixing families is MIXED_MODE', () => {
        const { mode, error } = modeOf(['breakdown', 'register', 'ledgerStatement']);
        expect(mode).toBeUndefined();
        expect(error).toEqual(expect.objectContaining({ code: 'MIXED_MODE' }));
    });

    it('half a register input is INCOMPLETE_REGISTER_INPUT, naming the missing side', () => {
        const registerOnly = modeOf(['register']);
        expect(registerOnly.error).toEqual(expect.objectContaining({ code: 'INCOMPLETE_REGISTER_INPUT' }));
        expect(registerOnly.error!.message).toMatch(/ledger/i);

        const statementOnly = modeOf(['ledgerStatement', 'ledgerStatement']);
        expect(statementOnly.error).toEqual(expect.objectContaining({ code: 'INCOMPLETE_REGISTER_INPUT' }));
        expect(statementOnly.error!.message).toMatch(/register/i);
    });

    it('nothing recognisable resolves to neither mode nor error (legacy path decides)', () => {
        expect(modeOf(['unknown', 'unknown'])).toEqual({});
    });
});

describe('ingest routing (GOAL-3 R1)', () => {
    it('ingests the register-family fixture: 16 postings + 11 cheques, continuous rows', async () => {
        const result = await ingest(readFileSync(FIXTURE), { filename: 'register-sample.xlsx' });
        expect(result.mode).toBe('register');
        expect(result.errors).toHaveLength(0);
        expect(result.postings).toHaveLength(16);
        expect(result.cheques).toHaveLength(11);

        // Row numbering is continuous and unique across the Credit + Debit sheets.
        const rowNumbers = result.postings.map((p) => p.rowNumber).sort((a, b) => a - b);
        expect(rowNumbers).toEqual(Array.from({ length: 16 }, (_, i) => i + 1));

        expect(result.summary.parsed).toBe(16);
        expect(result.summary.dataRows).toBe(27); // 12 credits + 4 debits + 11 register rows
        expect(result.summary.debitCount).toBe(4);
        expect(result.summary.creditCount).toBe(12);
        expect(result.summary.netFils).toBe(-2765500); // derived credit balance (signed)
        expect(result.summary.branches).toEqual(['001', '002']);
    });

    it('rejects a workbook mixing both families with MIXED_MODE', async () => {
        const buffer = await workbookBuffer([
            { name: 'Breakdown', rows: [BREAKDOWN_HEADER as unknown[], ['BH', '1', 'D2810085', '2023-01-08', '020050 X', 'BHD', 10, 'J1']] },
            { name: 'Sheet1', rows: [REGISTER_HEADER as unknown[]] },
            { name: 'Credit', rows: [STATEMENT_HEADER as unknown[]] },
        ]);
        const result = await ingest(buffer, { filename: 'mixed.xlsx' });
        expect(result.postings).toHaveLength(0);
        expect(result.errors).toEqual([expect.objectContaining({ code: 'MIXED_MODE' })]);
    });

    it('rejects a lone ledger-statement sheet with INCOMPLETE_REGISTER_INPUT', async () => {
        const buffer = await workbookBuffer([
            {
                name: 'Credit',
                rows: [
                    STATEMENT_HEADER as unknown[],
                    [new Date('2025-03-10T00:00:00Z'), new Date('2025-03-10T00:00:00Z'), '99801000', 5001, '001', 100, null, 2730],
                ],
            },
        ]);
        const result = await ingest(buffer, { filename: 'credit-only.xlsx' });
        expect(result.postings).toHaveLength(0);
        expect(result.errors).toEqual([expect.objectContaining({ code: 'INCOMPLETE_REGISTER_INPUT' })]);
    });

    it('breakdown workbooks keep the existing behaviour and report mode breakdown', async () => {
        const buffer = await workbookBuffer([
            {
                name: 'Breakdown',
                rows: [
                    BREAKDOWN_HEADER as unknown[],
                    ['BH', '1', 'D2810085', '2023-01-08', '020050 DEBIT POSTING', 'BHD', 10, 'J1'],
                ],
            },
            { name: 'Notes', rows: [['item', 'comment'], ['x', 'y']] },
        ]);
        const result = await ingest(buffer, { filename: 'breakdown.xlsx' });
        expect(result.mode).toBe('breakdown');
        expect(result.cheques).toBeUndefined();
        expect(result.postings).toHaveLength(1);
        expect(result.errors.filter((e) => e.code === 'SHEET_SKIPPED')).toHaveLength(1);
    });

    it('csv input stays on the breakdown path', async () => {
        const csv = [
            'entity,Branch Number,gl,Post Date,Log description,ccy,Amount (BHD),Journal Number',
            'BH,1,D2810085,2023-01-08,020050 DEBIT POSTING,BHD,10.000,J1',
        ].join('\n');
        const result = await ingest(Buffer.from(csv), { filename: 'plain.csv' });
        expect(result.mode).toBe('breakdown');
        expect(result.postings).toHaveLength(1);
    });
});

const VAT_HEADER = [
    'Transaction Date', 'Posting Date', 'Nostro/BGL Account', 'Journal Number', 'Account Name',
    'Transaction Description', 'Cheque Number', 'Transaction Credit Amount', 'Transaction Debit Amount',
    'Transaction Type', 'Teller', 'Branch', 'End Date EoD Balance', 'Previous EoD Balance',
];
const vatRow = (credit: unknown, debit: unknown, eod: number) => [
    new Date('2023-01-09T00:00:00Z'), new Date('2023-01-09T00:00:00Z'), '8828010400010000', 100483992,
    'INPUT VAT RECEIVABLE MUBASHER - BHD', 'NPB MISC DEP DR', '', credit, debit,
    '01-Financial', 'System', '00001-Main Branch', eod, 0,
];

describe('GOAL-8 statement mode routing', () => {
    it('a lone statement whose account is a statement GL ingests as mode statement', async () => {
        const buffer = await workbookBuffer([
            { name: 'Nostro and BGL Account Statemen', rows: [
                ['Nostro & BGL Account Statement'], [], VAT_HEADER,
                vatRow(0.5, null, 0.499), vatRow(null, -0.5, 0), // debit −0.5 → +0.5 engine
            ] },
        ]);
        const result = await ingest(buffer, { filename: 'vat.xlsx' });
        expect(result.mode).toBe('statement');
        expect(result.cheques).toBeUndefined();
        expect(result.postings).toHaveLength(2);
        expect(result.summary.netFils).toBe(0); // +500 (debit) − 500 (credit)
        expect(result.errors.filter(e => e.code === 'INCOMPLETE_REGISTER_INPUT')).toHaveLength(0);
    });

    it('a lone statement whose account is a REGISTER GL stays INCOMPLETE_REGISTER_INPUT', async () => {
        const buffer = await workbookBuffer([
            { name: 'Credit', rows: [VAT_HEADER, [
                new Date('2025-03-10T00:00:00Z'), new Date('2025-03-10T00:00:00Z'), '99801000', 5001,
                'MC PAYABLE', 'X', '', 100, null, '01', 'T', '001', 2730, 0,
            ]] },
        ]);
        const result = await ingest(buffer, { filename: 'mgr-half.xlsx' });
        expect(result.postings).toHaveLength(0);
        expect(result.errors).toEqual([expect.objectContaining({ code: 'INCOMPLETE_REGISTER_INPUT' })]);
    });

    it('a statement-schema CSV routes to statement mode (not breakdown)', async () => {
        const csv = [
            VAT_HEADER.join(','),
            ['2023-01-09', '2023-01-09', '8828010400010000', 'J1', 'INPUT VAT', 'NPB', '', '0.5', '', '01', 'T', '00001-Main Branch', '0.499', '0'].join(','),
        ].join('\n');
        const result = await ingest(Buffer.from(csv), { filename: 'vat.sanitized.csv' });
        expect(result.mode).toBe('statement');
        expect(result.postings).toHaveLength(1);
    });
});
