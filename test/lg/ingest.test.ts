import * as ExcelJS from 'exceljs';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { detectFormat, ingest, isLegacyXls, rowsFromCsv } from '../../src/lg/ingest';

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
