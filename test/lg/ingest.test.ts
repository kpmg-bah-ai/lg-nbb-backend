import * as ExcelJS from 'exceljs';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { detectFormat, ingest, ingestFiles, isLegacyXls, rowsFromCsv } from '../../src/lg/ingest';

const fixture = (name: string) => readFileSync(join(__dirname, '..', 'fixtures', 'lg', name));

describe('rowsFromCsv', () => {
    it('parses quoted fields containing commas', () => {
        const rows = rowsFromCsv('a,b,c\n1,"x, y",3\n');
        expect(rows).toHaveLength(2);
        expect(rows[1]).toEqual(['1', 'x, y', '3']);
    });
});

describe('detectFormat', () => {
    it('uses the filename extension', () => {
        expect(detectFormat(Buffer.from(''), 'breakdown.xlsx')).toBe('xlsx');
        expect(detectFormat(Buffer.from(''), 'breakdown.csv')).toBe('csv');
    });

    it('falls back to the PK zip magic bytes for xlsx', () => {
        expect(detectFormat(Buffer.from([0x50, 0x4b, 0x03, 0x04]))).toBe('xlsx');
        expect(detectFormat(Buffer.from('entity,gl\n'))).toBe('csv');
    });

    it('honours an explicit override', () => {
        expect(detectFormat(Buffer.from([0x50, 0x4b]), 'x.csv', 'csv')).toBe('csv');
    });
});

describe('ingest xlsx (F1) — real Date cells + "Amount\\r\\n(BHD)" header', () => {
    it('reads the balanced fixture into 6 balanced postings', async () => {
        const result = await ingest(fixture('balanced-sample.xlsx'), { filename: 'balanced-sample.xlsx' });
        expect(result.errors).toHaveLength(0);
        expect(result.postings).toHaveLength(6);
        expect(result.summary.netFils).toBe(0);
        expect(result.summary.debitCount).toBe(3);
        expect(result.summary.creditCount).toBe(3);
        expect(result.summary.currencies).toEqual(['BHD']);
        // Date cell coerced, amount header with embedded CR mapped, fils precision kept.
        expect(result.postings[0].postDate).toBe('2023-01-08');
        expect(result.postings.map((p) => p.amountBhdFils)).toContain(5590614);
    });

    it('surfaces a non-zero net for the unbalanced fixture (debit without a matching credit)', async () => {
        const result = await ingest(fixture('unbalanced-sample.xlsx'), { filename: 'unbalanced-sample.xlsx' });
        expect(result.postings).toHaveLength(3);
        expect(result.summary.netFils).toBe(5590614);
        expect(result.summary.debitCount).toBe(2);
        expect(result.summary.creditCount).toBe(1);
    });
});

describe('ingest csv (F1)', () => {
    it('reads the balanced csv fixture to the same result as the xlsx', async () => {
        const result = await ingest(fixture('balanced-sample.csv'), { filename: 'balanced-sample.csv' });
        expect(result.errors).toHaveLength(0);
        expect(result.postings).toHaveLength(6);
        expect(result.summary.netFils).toBe(0);
        expect(result.postings[0].postDate).toBe('2023-01-08');
    });

    it('finds a header that is not on the first row (title/preamble rows)', async () => {
        const csv = [
            'GL Transaction Breakdown — generated 2026-06-30',
            '',
            'entity,Branch Number,gl,Post Date,Log description,ccy,Amount (BHD),Journal Number',
            'BH,1,D2810085,2023-01-08,020050 DEBIT POSTING,BHD,10.000,J1',
        ].join('\n');
        const result = await ingest(Buffer.from(csv), { filename: 'preamble.csv' });
        expect(result.errors).toHaveLength(0);
        expect(result.postings).toHaveLength(1);
        expect(result.postings[0].amountBhdFils).toBe(10000);
    });
});

describe('ingest legacy .xls (F1)', () => {
    const oleMagic = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00, 0x00, 0x00, 0x00]);

    it('detects the OLE compound-file magic', () => {
        expect(isLegacyXls(oleMagic)).toBe(true);
        expect(isLegacyXls(fixture('balanced-sample.xlsx'))).toBe(false);
    });

    it('rejects .xls with UNSUPPORTED_FORMAT instead of mangling it as csv', async () => {
        const result = await ingest(oleMagic, { filename: 'legacy.xls' });
        expect(result.postings).toHaveLength(0);
        expect(result.errors).toHaveLength(1);
        expect(result.errors[0].code).toBe('UNSUPPORTED_FORMAT');
    });
});

describe('ingest multi-sheet xlsx (F1, GOAL.md §2.3)', () => {
    const HEADER = ['entity', 'Branch Number', 'gl', 'Post Date', 'Log description', 'ccy', 'Amount (BHD)', 'Journal Number'];
    const dataRow = (amount: number, journal: string) => ['BH', '1', 'D2810085', '2023-01-08', '020050 DEBIT POSTING', 'BHD', amount, journal];

    async function workbookBuffer(sheets: { name: string; rows: unknown[][] }[]): Promise<Buffer> {
        const workbook = new ExcelJS.Workbook();
        for (const sheet of sheets) {
            const ws = workbook.addWorksheet(sheet.name);
            ws.addRows(sheet.rows);
        }
        return Buffer.from(await workbook.xlsx.writeBuffer());
    }

    it('skips a sheet without a breakdown header and keeps the valid one', async () => {
        const buffer = await workbookBuffer([
            { name: 'Breakdown', rows: [HEADER, dataRow(10, 'J1'), dataRow(20, 'J2')] },
            { name: 'Mismatched', rows: [['item', 'comment'], ['x', 'y']] },
        ]);
        const result = await ingest(buffer, { filename: 'two-sheets.xlsx' });
        expect(result.postings).toHaveLength(2);
        const skipped = result.errors.filter((e) => e.code === 'SHEET_SKIPPED');
        expect(skipped).toHaveLength(1);
        expect(result.postings.every((p) => p.sheet !== undefined && p.sheet !== '')).toBe(true);
        expect(result.postings[0].sheet).not.toBe(skipped[0].sheet);
    });

    it('appends rows from every valid sheet with continuous row numbering', async () => {
        const buffer = await workbookBuffer([
            { name: 'January', rows: [HEADER, dataRow(10, 'J1'), dataRow(20, 'J2')] },
            { name: 'February', rows: [HEADER, dataRow(30, 'J3')] },
        ]);
        const result = await ingest(buffer, { filename: 'two-valid-sheets.xlsx' });
        expect(result.errors).toHaveLength(0);
        expect(result.postings).toHaveLength(3);
        expect(result.postings.map((p) => p.rowNumber)).toEqual([1, 2, 3]);
        expect(result.summary.dataRows).toBe(3);
        expect(result.summary.netFils).toBe(60000);
    });

    it('rejects a workbook where no sheet has a recognisable header', async () => {
        const buffer = await workbookBuffer([
            { name: 'A', rows: [['foo', 'bar'], [1, 2]] },
            { name: 'B', rows: [['baz'], [3]] },
        ]);
        const result = await ingest(buffer, { filename: 'no-headers.xlsx' });
        expect(result.postings).toHaveLength(0);
        expect(result.errors.some((e) => e.code === 'MISSING_HEADER')).toBe(true);
    });
});

/**
 * Multi-file upload: every file's worksheets pool into ONE run before role
 * detection, so a register family split across files (ledger extract in one,
 * cheque register in another) resolves — while mixed families still reject.
 * Sheet names gain a "filename › sheet" prefix for per-file provenance.
 */
describe('ingestFiles (multi-file upload)', () => {
    const HEADER = ['entity', 'Branch Number', 'gl', 'Post Date', 'Log description', 'ccy', 'Amount (BHD)', 'Journal Number'];
    const dataRow = (amount: number, journal: string) => ['BH', '1', 'D2810085', '2023-01-08', '020050 DEBIT POSTING', 'BHD', amount, journal];

    const STATEMENT_HEADER = [
        'Transaction Date',
        'Posting Date',
        'Nostro/BGL Account',
        'Journal Number',
        'Transaction Credit Amount',
        'Transaction Debit Amount',
        'Branch',
        'End Date EoD Balance',
        'Detailed Description',
    ];
    const REGISTER_HEADER = [
        'c2_instr_type',
        'c6_chq_no',
        'c9_amount',
        'c16_issued_date',
        'c19_issued_jrnl_no',
        'c25_matchd_post_dt',
        'c27_matchd_jrnl_no',
    ];

    async function workbookBuffer(sheets: { name: string; rows: unknown[][] }[]): Promise<Buffer> {
        const workbook = new ExcelJS.Workbook();
        for (const sheet of sheets) {
            const ws = workbook.addWorksheet(sheet.name);
            ws.addRows(sheet.rows);
        }
        return Buffer.from(await workbook.xlsx.writeBuffer());
    }

    it('pools breakdown sheets from multiple files into one run with continuous row numbering', async () => {
        const jan = await workbookBuffer([{ name: 'January', rows: [HEADER, dataRow(10, 'J1'), dataRow(20, 'J2')] }]);
        const feb = await workbookBuffer([{ name: 'February', rows: [HEADER, dataRow(30, 'J3')] }]);
        const result = await ingestFiles([
            { buffer: jan, filename: 'january.xlsx' },
            { buffer: feb, filename: 'february.xlsx' },
        ]);
        expect(result.mode).toBe('breakdown');
        expect(result.errors).toHaveLength(0);
        expect(result.postings).toHaveLength(3);
        expect(result.postings.map((p) => p.rowNumber)).toEqual([1, 2, 3]);
        expect(result.summary.dataRows).toBe(3);
        expect(result.summary.netFils).toBe(60000);
        // Per-file provenance: sheet names carry the source filename.
        expect(result.postings[0].sheet).toBe('january.xlsx › January');
        expect(result.postings[2].sheet).toBe('february.xlsx › February');
    });

    it('mixes xlsx and csv files in one upload', async () => {
        const xlsx = await workbookBuffer([{ name: 'Breakdown', rows: [HEADER, dataRow(10, 'J1')] }]);
        const csv = Buffer.from(
            [HEADER.join(','), 'BH,1,D2810085,2023-01-08,020050 DEBIT POSTING,BHD,20.000,J2'].join('\n')
        );
        const result = await ingestFiles([
            { buffer: xlsx, filename: 'part1.xlsx' },
            { buffer: csv, filename: 'part2.csv' },
        ]);
        expect(result.errors).toHaveLength(0);
        expect(result.postings).toHaveLength(2);
        expect(result.postings[1].sheet).toBe('part2.csv');
        expect(result.summary.netFils).toBe(30000);
    });

    it('resolves a register family split across two files (ledger + register)', async () => {
        const ledger = await workbookBuffer([
            {
                name: 'Credit',
                rows: [STATEMENT_HEADER, ['2025-03-10', '2025-03-11', '99801000', 'J5001', 100, null, '001', 2730, '']],
            },
            {
                name: 'Debit',
                rows: [STATEMENT_HEADER, ['2025-04-01', null, '99801000', 'J6001', null, -100, '001', 2730, '']],
            },
        ]);
        const register = await workbookBuffer([
            {
                name: 'Sheet1',
                rows: [REGISTER_HEADER, ['PO', 1001, 100, '2025-03-10', 'J5001', '2025-04-01', 'J6001']],
            },
        ]);
        const result = await ingestFiles([
            { buffer: ledger, filename: 'ledger.xlsx' },
            { buffer: register, filename: 'register.xlsx' },
        ]);
        expect(result.mode).toBe('register');
        expect(result.errors.filter((e) => e.code === 'INCOMPLETE_REGISTER_INPUT')).toHaveLength(0);
        expect(result.postings).toHaveLength(2);
        expect(result.cheques).toHaveLength(1);
        expect(result.postings[0].sheet).toBe('ledger.xlsx › Credit');
        expect(result.cheques![0].sheet).toBe('register.xlsx › Sheet1');
        expect(result.summary.dataRows).toBe(3); // 2 ledger rows + 1 register row
    });

    it('still rejects mixed families across files with MIXED_MODE', async () => {
        const result = await ingestFiles([
            { buffer: fixture('balanced-sample.xlsx'), filename: 'breakdown.xlsx' },
            { buffer: fixture('register-sample.xlsx'), filename: 'register.xlsx' },
        ]);
        expect(result.postings).toHaveLength(0);
        expect(result.errors).toEqual([expect.objectContaining({ code: 'MIXED_MODE' })]);
    });

    it('a single file behaves exactly like ingest() — no filename prefix on sheets', async () => {
        const single = await ingestFiles([{ buffer: fixture('balanced-sample.xlsx'), filename: 'balanced-sample.xlsx' }]);
        const direct = await ingest(fixture('balanced-sample.xlsx'), { filename: 'balanced-sample.xlsx' });
        expect(single).toEqual(direct);
    });

    it('rejects the whole upload when any file is legacy .xls, naming the file', async () => {
        const ole = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
        const result = await ingestFiles([
            { buffer: fixture('balanced-sample.xlsx'), filename: 'ok.xlsx' },
            { buffer: ole, filename: 'legacy.xls' },
        ]);
        expect(result.postings).toHaveLength(0);
        expect(result.errors).toHaveLength(1);
        expect(result.errors[0].code).toBe('UNSUPPORTED_FORMAT');
        expect(result.errors[0].message).toContain('legacy.xls');
    });

    it('an empty file list is EMPTY_INPUT', async () => {
        const result = await ingestFiles([]);
        expect(result.postings).toHaveLength(0);
        expect(result.errors[0].code).toBe('EMPTY_INPUT');
    });
});
