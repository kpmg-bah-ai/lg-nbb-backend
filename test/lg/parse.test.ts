import { RawRow } from '../../src/shared/models';
import {
    coerceDate,
    deriveDirection,
    excelColumn,
    extractLogCode,
    mapHeaders,
    normalizeHeader,
    normalizeRow,
    parseAmountToFils,
    parseRows,
    parseSheets,
} from '../../src/lg/parse';

// A trimmed header mirroring the real breakdown, including the CR/LF escape that
// appears inside the amount headers and the stray "?column?" we must ignore.
const HEADER: RawRow = [
    'entity',
    'Branch Number',
    'gl',
    'Post Date',
    'Log description',
    'ccy',
    'Amount_x000D_\n(BHD)',
    'Journal Number',
    '?column?',
];

function dataRow(amount: unknown, logdesc: string, journal: string, postDate = '2023-01-08'): RawRow {
    return ['BH', '1', 'D2810085', postDate, logdesc, 'BHD', amount as never, journal, 'x'];
}

describe('normalizeHeader', () => {
    it('strips the _x000D_ escape, whitespace and case', () => {
        expect(normalizeHeader('Amount_x000D_\n(BHD)')).toBe('amount(bhd)');
        expect(normalizeHeader('Branch Number')).toBe('branchnumber');
        expect(normalizeHeader('gl_desc')).toBe('gldesc');
    });
});

describe('mapHeaders', () => {
    it('maps known columns and ignores ?column?', () => {
        const { columns, errors } = mapHeaders(HEADER);
        expect(errors).toHaveLength(0);
        expect(columns.amountBhd).toBe(6);
        expect(columns.journalNumber).toBe(7);
    });

    it('reports every missing required column', () => {
        const { errors } = mapHeaders(['entity', 'gl'] as RawRow);
        const missing = errors.map((e) => e.field).sort();
        expect(errors.every((e) => e.code === 'MISSING_HEADER')).toBe(true);
        expect(missing).toEqual(['amountBhd', 'branchNumber', 'currency', 'journalNumber', 'logDescription', 'postDate']);
    });
});

describe('parseAmountToFils (BHD = 3dp / fils, no floats)', () => {
    it('keeps fils precision on fractional amounts', () => {
        expect(parseAmountToFils(5590.614)).toBe(5590614);
        expect(parseAmountToFils(0.555)).toBe(555);
        expect(parseAmountToFils(-400)).toBe(-400000);
    });

    it('parses numeric strings, thousands separators and parentheses-negatives', () => {
        expect(parseAmountToFils('1,553')).toBe(1553000);
        expect(parseAmountToFils('(245)')).toBe(-245000);
    });

    it('returns undefined for blank / non-numeric', () => {
        expect(parseAmountToFils('')).toBeUndefined();
        expect(parseAmountToFils(undefined)).toBeUndefined();
        expect(parseAmountToFils('n/a')).toBeUndefined();
    });
});

describe('coerceDate', () => {
    it('handles Date objects, ISO strings and Excel serials', () => {
        expect(coerceDate(new Date(Date.UTC(2023, 0, 8)))).toBe('2023-01-08');
        expect(coerceDate('2026-06-09T00:00:00Z')).toBe('2026-06-09');
        expect(coerceDate(44934)).toBe('2023-01-08'); // Excel serial for 2023-01-08
    });

    it('returns undefined for junk', () => {
        expect(coerceDate('not a date')).toBeUndefined();
        expect(coerceDate(undefined)).toBeUndefined();
    });
});

describe('extractLogCode / deriveDirection', () => {
    it('pulls the leading 6-digit posting code', () => {
        expect(extractLogCode('020050 DEBIT POSTING')).toBe('020050');
        expect(extractLogCode('BAL TRF')).toBeUndefined();
    });

    it('derives direction from sign, falling back to log code on zero', () => {
        expect(deriveDirection(400000)).toBe('debit');
        expect(deriveDirection(-400000)).toBe('credit');
        expect(deriveDirection(0, '020050')).toBe('debit');
        expect(deriveDirection(0, '020030')).toBe('credit');
        expect(deriveDirection(0)).toBeUndefined();
    });
});

describe('normalizeRow', () => {
    const { columns } = mapHeaders(HEADER);

    it('normalises a valid row into a posting (fils + direction + traceable row number)', () => {
        const { posting, errors } = normalizeRow(dataRow(-5590.614, '020030 BGL CR POSTING', '100035429'), columns, 3);
        expect(errors).toHaveLength(0);
        expect(posting).toMatchObject({
            branchNumber: '1',
            gl: 'D2810085',
            amountBhdFils: -5590614,
            amountBhd: -5590.614,
            direction: 'credit',
            logCode: '020030',
            journalNumber: '100035429',
            postDate: '2023-01-08',
            rowNumber: 3,
        });
    });

    it('flags a bad amount and drops the posting', () => {
        const { posting, errors } = normalizeRow(dataRow('oops', '020050 DEBIT POSTING', 'J1'), columns, 1);
        expect(posting).toBeUndefined();
        expect(errors).toContainEqual(expect.objectContaining({ code: 'BAD_AMOUNT', field: 'amountBhd', row: 1 }));
    });

    it('flags a missing required value', () => {
        const { posting, errors } = normalizeRow(dataRow(100, '020050 DEBIT POSTING', ''), columns, 2);
        expect(posting).toBeUndefined();
        expect(errors).toContainEqual(expect.objectContaining({ code: 'MISSING_FIELD', field: 'journalNumber', row: 2 }));
    });
});

describe('parseRows', () => {
    it('parses a balanced set: net fils 0, matched debit/credit counts', () => {
        const rows: RawRow[] = [
            HEADER,
            dataRow(-400, '020030 BGL CR POSTING', 'J1'),
            dataRow(400, '020050 DEBIT POSTING', 'J2'),
            dataRow(-0.555, '020030 BGL CR POSTING', 'J3'),
            dataRow(0.555, '020050 DEBIT POSTING', 'J4'),
        ];
        const result = parseRows(rows);
        expect(result.errors).toHaveLength(0);
        expect(result.postings).toHaveLength(4);
        expect(result.summary).toMatchObject({
            dataRows: 4,
            parsed: 4,
            debitCount: 2,
            creditCount: 2,
            netFils: 0,
            currencies: ['BHD'],
            branches: ['1'],
        });
    });

    it('reports a non-zero net when a debit has no matching credit', () => {
        const rows: RawRow[] = [HEADER, dataRow(-400, '020030 BGL CR POSTING', 'J1'), dataRow(400, '020050 DEBIT POSTING', 'J2'), dataRow(5590.614, '020050 DEBIT POSTING', 'J3')];
        const result = parseRows(rows);
        expect(result.summary.netFils).toBe(5590614);
        expect(result.summary.debitCount).toBe(2);
        expect(result.summary.creditCount).toBe(1);
    });

    it('short-circuits on a missing header', () => {
        const result = parseRows([['entity', 'gl'] as RawRow, dataRow(1, 'x', 'J1')]);
        expect(result.postings).toHaveLength(0);
        expect(result.errors.some((e) => e.code === 'MISSING_HEADER')).toBe(true);
    });

    it('reports EMPTY_INPUT for no rows and skips blank rows', () => {
        expect(parseRows([]).errors).toContainEqual(expect.objectContaining({ code: 'EMPTY_INPUT' }));
        const withBlank = parseRows([HEADER, [] as RawRow, dataRow(400, '020050 DEBIT POSTING', 'J1')]);
        expect(withBlank.summary.dataRows).toBe(1);
        expect(withBlank.postings).toHaveLength(1);
    });
});

/**
 * Wrong-value tracking: a non-empty cell that fails its column's format is
 * saved on the error (raw value + spreadsheet column + header). Wrong values
 * in optional columns never drop the row — it parses, and the wrongness is
 * tracked with rowParsed. Empty cells are missing, never "wrong".
 */
describe('wrong-value tracking', () => {
    const { columns } = mapHeaders(HEADER);

    it('excelColumn turns 0-based indexes into spreadsheet letters', () => {
        expect(excelColumn(0)).toBe('A');
        expect(excelColumn(6)).toBe('G');
        expect(excelColumn(25)).toBe('Z');
        expect(excelColumn(26)).toBe('AA');
    });

    it('BAD_AMOUNT saves the wrong raw value and its column', () => {
        const { posting, errors } = normalizeRow(dataRow('oops', '020050 DEBIT POSTING', 'J1'), columns, 1, HEADER);
        expect(posting).toBeUndefined();
        expect(errors).toContainEqual(
            expect.objectContaining({
                code: 'BAD_AMOUNT',
                value: 'oops',
                column: 'G',
                columnHeader: 'Amount_x000D_\n(BHD)',
            })
        );
    });

    it('BAD_DATE saves the wrong raw value and its column', () => {
        const { errors } = normalizeRow(dataRow(100, '020050 DEBIT POSTING', 'J1', 'journal 987'), columns, 1, HEADER);
        expect(errors).toContainEqual(
            expect.objectContaining({
                code: 'BAD_DATE',
                field: 'postDate',
                value: 'journal 987',
                column: 'D',
                columnHeader: 'Post Date',
            })
        );
    });

    it('ZERO_AMOUNT saves the zero and its column', () => {
        const { errors } = normalizeRow(dataRow(0, 'BAL TRF', 'J9'), columns, 1, HEADER);
        expect(errors).toContainEqual(
            expect.objectContaining({ code: 'ZERO_AMOUNT', value: '0', column: 'G' })
        );
    });

    it('MISSING_FIELD names the column but carries no value (empty is missing, not wrong)', () => {
        const { errors } = normalizeRow(dataRow(100, '020050 DEBIT POSTING', ''), columns, 2, HEADER);
        const err = errors.find((e) => e.code === 'MISSING_FIELD');
        expect(err).toMatchObject({ field: 'journalNumber', column: 'H', columnHeader: 'Journal Number' });
        expect(err?.value).toBeUndefined();
    });

    it('a wrong optional value (Vale Date) is tracked but the row STILL parses', () => {
        const header: RawRow = [...HEADER, 'Vale Date'];
        const { columns: cols } = mapHeaders(header);
        const raw: RawRow = [...dataRow(-400, '020030 BGL CR POSTING', 'J1'), 'not-a-date'];
        const { posting, errors } = normalizeRow(raw, cols, 1, header);
        expect(posting).toBeDefined();
        expect(posting?.valueDate).toBeUndefined();
        expect(errors).toContainEqual(
            expect.objectContaining({
                code: 'BAD_DATE',
                field: 'valueDate',
                value: 'not-a-date',
                column: 'J',
                columnHeader: 'Vale Date',
                rowParsed: true,
            })
        );
    });

    it('an empty optional value is not wrong — no tracking entry', () => {
        const header: RawRow = [...HEADER, 'Vale Date'];
        const { columns: cols } = mapHeaders(header);
        const { posting, errors } = normalizeRow([...dataRow(-400, '020030 BGL CR POSTING', 'J1'), ''], cols, 1, header);
        expect(posting).toBeDefined();
        expect(errors).toHaveLength(0);
    });

    it('a wrong FCY amount is tracked but never drops the row', () => {
        const header: RawRow = [...HEADER, 'Amount (FCY)'];
        const { columns: cols } = mapHeaders(header);
        const { posting, errors } = normalizeRow([...dataRow(-400, '020030 BGL CR POSTING', 'J1'), '12x'], cols, 1, header);
        expect(posting).toBeDefined();
        expect(posting?.amountFcyFils).toBeUndefined();
        expect(errors).toContainEqual(
            expect.objectContaining({
                code: 'BAD_AMOUNT',
                field: 'amountFcy',
                value: '12x',
                column: 'J',
                rowParsed: true,
            })
        );
    });

    it('parseSheets carries the tracking through with the sheet name', () => {
        const result = parseSheets([{ name: 'Breakdown', rows: [HEADER, dataRow('bad', '020050 DEBIT POSTING', 'J1')] }]);
        expect(result.errors).toContainEqual(
            expect.objectContaining({ code: 'BAD_AMOUNT', sheet: 'Breakdown', value: 'bad', column: 'G' })
        );
    });
});
