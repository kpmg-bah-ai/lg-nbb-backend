/**
 * GOAL-3 R7 — register-mode exception classifier (src/lg/registerExceptions.ts).
 *
 * Every outstanding item becomes exactly one exception (nothing dropped), with
 * register-aware reasons; on top come informational KEY_COLLISION annotations
 * and the run-level EXTRACT_GAP (GOAL-3 §4.6).
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { ingest } from '../../src/lg/ingest';
import { matchRegister } from '../../src/lg/registerMatch';
import { classifyRegisterExceptions } from '../../src/lg/registerExceptions';
import { ParsedPosting, RegisterCheque } from '../../src/shared/models';

const FIXTURE = join(__dirname, '..', 'fixtures', 'lg', 'register-sample.xlsx');
const AS_OF = '2026-02-03';

let rowCounter = 0;

function posting(overrides: Partial<ParsedPosting> & { amountBhdFils: number }): ParsedPosting {
    const fils = overrides.amountBhdFils;
    const date = overrides.postDate ?? '2025-06-01';
    return {
        entity: '',
        branchNumber: '001',
        gl: '99801000',
        postDate: date,
        logDescription: '',
        currency: 'BHD',
        amountBhd: fils / 1000,
        direction: fils >= 0 ? 'debit' : 'credit',
        journalNumber: 'J1',
        rowNumber: ++rowCounter,
        transactionDate: date,
        ...overrides,
    };
}

function cheque(overrides: Partial<RegisterCheque> & { amountFils: number }): RegisterCheque {
    return { chequeNumber: '9000', status: '01', opsPaid: false, rowNumber: ++rowCounter, ...overrides };
}

async function fixtureExceptions() {
    const result = await ingest(readFileSync(FIXTURE), { filename: 'register-sample.xlsx' });
    const match = matchRegister(result.postings, result.cheques!, { asOf: AS_OF });
    return { match, ...classifyRegisterExceptions(match, -35500) };
}

describe('classifyRegisterExceptions — fixture', () => {
    test('classifies every designed case exactly once', async () => {
        const { summary } = await fixtureExceptions();
        expect(summary.byReason).toEqual({
            UNMATCHED_CREDIT: 4, // statement cheques 1002/1006/1009/1011
            NON_ISSUANCE_CREDIT: 2, // take-on + COR CHQ Redeep
            REGISTER_LAG_OPS_PAID: 1, // 1004
            REGISTER_PAID_NO_LEDGER_DEBIT: 1, // 1010
            UNMATCHED_LEDGER_DEBIT: 1, // stray D4
            KEY_COLLISION: 2, // 1005 + 1006 annotations
            EXTRACT_GAP: 1,
        });
        expect(summary.total).toBe(12);
    });

    test('exceptions ⊇ outstanding: every outstanding row survives classification', async () => {
        const { match, exceptions } = await fixtureExceptions();
        const exceptionRows = new Set(exceptions.map((e) => `${e.sheet ?? ''}:${e.rowNumber}`));
        for (const item of match.outstanding) {
            expect(exceptionRows.has(`${item.sheet ?? ''}:${item.rowNumber}`)).toBe(true);
        }
        // Amounts pass through untouched — Σ outstanding is preserved exactly.
        const outstandingFils = match.outstanding.reduce((s, o) => s + o.outstandingFils, 0);
        const nonInformational = exceptions.filter((e) => e.reason !== 'KEY_COLLISION' && e.reason !== 'EXTRACT_GAP');
        expect(nonInformational.reduce((s, e) => s + e.outstandingFils, 0)).toBe(outstandingFils);
    });

    test('messages are reviewer-readable and carry the evidence', async () => {
        const { exceptions } = await fixtureExceptions();
        const opsLag = exceptions.find((e) => e.reason === 'REGISTER_LAG_OPS_PAID')!;
        expect(opsLag.message).toMatch(/PAID/);
        expect(opsLag.message).toMatch(/7001/); // the ops journal
        expect(opsLag.message).toMatch(/register/i);

        const nonIssuance = exceptions.find((e) => e.journalNumber === '999999999')!;
        expect(nonIssuance.reason).toBe('NON_ISSUANCE_CREDIT');
        expect(nonIssuance.message).toMatch(/issuance/i);

        const regPaid = exceptions.find((e) => e.reason === 'REGISTER_PAID_NO_LEDGER_DEBIT')!;
        expect(regPaid.message).toMatch(/2026-01-25/); // the register matched date
        expect(regPaid.message).toMatch(/6010/); // the register matched journal

        const gap = exceptions.find((e) => e.reason === 'EXTRACT_GAP')!;
        expect(gap.message).toMatch(/35\.500/);
        expect(gap.message).toMatch(/derived/i);

        const collision = exceptions.filter((e) => e.reason === 'KEY_COLLISION');
        expect(collision.map((e) => e.message.includes('one-for-one')).every(Boolean)).toBe(true);
    });

    test('outstanding statement cheques keep their base reason with cheque context', async () => {
        const { exceptions } = await fixtureExceptions();
        const statementLines = exceptions.filter((e) => e.reason === 'UNMATCHED_CREDIT');
        expect(statementLines).toHaveLength(4);
        for (const line of statementLines) {
            expect(line.message).toMatch(/cheque/i);
        }
    });
});

describe('classifyRegisterExceptions — micro-cases', () => {
    test('a partially allocated batch debit is UNRESOLVED_BATCH_DEBIT', () => {
        const credit = posting({ amountBhdFils: -120000, transactionDate: '2025-06-01', journalNumber: '5100' });
        const batchDebit = posting({
            amountBhdFils: 200000,
            transactionDate: '2025-07-01',
            journalNumber: '8100',
            detailedDescription: 'DEBIT POSTING-20-Ref.# 5100,Ref.# 5199',
        });
        const chq = cheque({
            amountFils: 120000,
            issuedDate: '2025-06-01',
            issuedPostDate: '2025-06-01',
            issuedJournal: '5100',
        });
        const match = matchRegister([credit, batchDebit], [chq], { asOf: '2026-01-01' });
        const { exceptions } = classifyRegisterExceptions(match, undefined);
        const batch = exceptions.find((e) => e.journalNumber === '8100')!;
        expect(batch.reason).toBe('UNRESOLVED_BATCH_DEBIT');
        expect(batch.message).toMatch(/5199/); // the unresolved ref is named
        expect(batch.outstandingFils).toBe(80000);
    });

    test('a manually dispositioned debit keeps its note in the message', () => {
        const debit = posting({
            amountBhdFils: 10000,
            transactionDate: '2025-07-01',
            journalNumber: '9998',
            reconciledNote: 'Datafix entry - Not manager cheque transaction',
        });
        const match = matchRegister([debit], [], { asOf: '2026-01-01' });
        const { exceptions } = classifyRegisterExceptions(match, undefined);
        expect(exceptions[0].reason).toBe('UNMATCHED_LEDGER_DEBIT');
        expect(exceptions[0].message).toMatch(/Datafix entry/);
    });

    test('no gap means no EXTRACT_GAP entry', () => {
        const match = matchRegister([], [], { asOf: '2026-01-01' });
        const { exceptions, summary } = classifyRegisterExceptions(match, 0);
        expect(exceptions).toHaveLength(0);
        expect(summary.total).toBe(0);
    });
});
