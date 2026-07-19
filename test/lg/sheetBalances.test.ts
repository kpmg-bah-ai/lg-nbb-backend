import { computeSheetBalances } from '../../src/lg/sheetBalances';
import { IngestResult } from '../../src/lg/ingest';
import { ParsedPosting, RegisterCheque, ParseError } from '../../src/shared/models';
import { makePosting } from './helpers';

function ingestResult(over: Partial<IngestResult>): IngestResult {
    return {
        mode: 'register',
        postings: [],
        cheques: [],
        errors: [],
        summary: { dataRows: 0, parsed: 0, debitCount: 0, creditCount: 0, netFils: 0, currencies: [], branches: [] },
        ...over,
    };
}

function cheque(over: Partial<RegisterCheque> & { amountFils: number }): RegisterCheque {
    return { opsPaid: false, rowNumber: 1, ...over };
}

describe('computeSheetBalances (GOAL-5 per-sheet balance reference)', () => {
    it('totals credits and debits per worksheet with a signed net', () => {
        const postings: ParsedPosting[] = [
            makePosting({ amountBhdFils: -1000, sheet: 'Credit' }),
            makePosting({ amountBhdFils: -500, sheet: 'Credit' }),
            makePosting({ amountBhdFils: 2000, sheet: 'Debit' }),
        ];
        const balances = computeSheetBalances(ingestResult({ postings }));

        const credit = balances.find((b) => b.sheet === 'Credit')!;
        expect(credit.role).toBe('ledger');
        expect(credit.creditCount).toBe(2);
        expect(credit.debitCount).toBe(0);
        expect(credit.creditFils).toBe(1500);
        expect(credit.netFils).toBe(-1500); // both credits

        const debit = balances.find((b) => b.sheet === 'Debit')!;
        expect(debit.debitFils).toBe(2000);
        expect(debit.netFils).toBe(2000);
    });

    it('captures the stated End Date EoD balance from a ledger sheet final-day rows', () => {
        const postings: ParsedPosting[] = [
            makePosting({ amountBhdFils: -1000, sheet: 'Credit', postDate: '2026-01-01', statedEodFils: 4000 }),
            makePosting({ amountBhdFils: -3000, sheet: 'Credit', postDate: '2026-02-03', statedEodFils: 7000 }),
        ];
        const [credit] = computeSheetBalances(ingestResult({ postings }));
        // The stated figure is a credit balance magnitude → negated into engine-signed fils.
        expect(credit.statedEodFils).toBe(-7000);
        expect(credit.basis).toMatch(/End Date EoD/i);
    });

    it('summarises a register sheet by cheque count and amount', () => {
        const cheques: RegisterCheque[] = [
            cheque({ amountFils: 1000, sheet: 'Sheet1', rowNumber: 1 }),
            cheque({ amountFils: 2500, sheet: 'Sheet1', rowNumber: 2 }),
        ];
        const [reg] = computeSheetBalances(ingestResult({ postings: [], cheques }));
        expect(reg.role).toBe('register');
        expect(reg.chequeCount).toBe(2);
        expect(reg.chequeFils).toBe(3500);
        expect(reg.parsedRows).toBe(0);
    });

    it('lists skipped worksheets with a zero balance and a reason', () => {
        const errors: ParseError[] = [
            { code: 'SHEET_SKIPPED', sheet: 'Notes', message: 'Worksheet "Notes" matches no schema — skipped' },
        ];
        const balances = computeSheetBalances(ingestResult({ errors }));
        const notes = balances.find((b) => b.sheet === 'Notes')!;
        expect(notes.role).toBe('skipped');
        expect(notes.creditFils).toBe(0);
        expect(notes.basis).toMatch(/skipped/i);
    });

    it('is deterministic: sheets are ordered ledger → register → skipped, then by name', () => {
        const postings = [makePosting({ amountBhdFils: -1, sheet: 'Debit' })];
        const cheques = [cheque({ amountFils: 1, sheet: 'Sheet1' })];
        const errors: ParseError[] = [{ code: 'SHEET_SKIPPED', sheet: 'Notes', message: 'skipped' }];
        const balances = computeSheetBalances(ingestResult({ postings, cheques, errors }));
        expect(balances.map((b) => b.sheet)).toEqual(['Debit', 'Sheet1', 'Notes']);
    });

    it('statement mode uses the ledger role and captures the stated EoD', () => {
        const balances = computeSheetBalances({
            mode: 'statement',
            postings: [
                makePosting({ amountBhdFils: 500, sheet: 'VAT', statedEodFils: -1000, postDate: '2023-02-01' }),
                makePosting({ amountBhdFils: 500, sheet: 'VAT', statedEodFils: -1000, postDate: '2023-02-01' }),
            ],
            errors: [],
        });
        expect(balances[0].role).toBe('ledger');
        expect(balances[0].netFils).toBe(1000);       // Σ debit − credit
        expect(balances[0].statedEodFils).toBe(1000); // −(−1000) engine-signed
    });
});
