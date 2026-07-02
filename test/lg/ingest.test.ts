import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { detectFormat, ingest, rowsFromCsv } from '../../src/lg/ingest';

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
});
