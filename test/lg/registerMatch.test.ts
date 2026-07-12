/**
 * GOAL-3 R4 — two-legged GL↔register matcher (src/lg/registerMatch.ts).
 *
 * Tuple key (transaction date, journal, |amount fils|) with MULTISET semantics:
 * colliding keys pair off one-for-one FIFO, never via lookup (file §3 caveat).
 * A MatchedSet only forms when both legs are in the ledger window, so every
 * set nets to zero and the engine invariant survives verbatim:
 *   Σ signed outstanding = Σ signed included postings.
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { ingest } from '../../src/lg/ingest';
import { matchRegister } from '../../src/lg/registerMatch';
import { ChequeOutcome, ParsedPosting, RegisterCheque } from '../../src/shared/models';

const FIXTURE = join(__dirname, '..', 'fixtures', 'lg', 'register-sample.xlsx');
const AS_OF = '2026-02-03';

let rowCounter = 0;

/** A ledger-statement posting: credits negative, debits positive (engine convention). */
function posting(overrides: Partial<ParsedPosting> & { amountBhdFils: number }): ParsedPosting {
    const fils = overrides.amountBhdFils;
    const debit = fils >= 0;
    const date = overrides.transactionDate ?? '2025-06-01';
    return {
        entity: '',
        branchNumber: '001',
        gl: '99801000',
        postDate: date,
        logDescription: '',
        currency: 'BHD',
        amountBhd: fils / 1000,
        direction: debit ? 'debit' : 'credit',
        journalNumber: 'J1',
        rowNumber: ++rowCounter,
        transactionDate: date,
        ...overrides,
    };
}

function cheque(overrides: Partial<RegisterCheque> & { amountFils: number }): RegisterCheque {
    return {
        chequeNumber: '9000',
        status: '01',
        issuedBranch: '001',
        opsPaid: false,
        rowNumber: ++rowCounter,
        ...overrides,
    };
}

async function fixtureRun() {
    const result = await ingest(readFileSync(FIXTURE), { filename: 'register-sample.xlsx' });
    return matchRegister(result.postings, result.cheques!, { asOf: AS_OF });
}

describe('matchRegister — fixture (passes 1–3)', () => {
    test('forms zero-net matched sets for cheques 1001 and 1005', async () => {
        const { matchedSets } = await fixtureRun();
        expect(matchedSets).toHaveLength(2);
        const byChq = new Map(matchedSets.map((s) => [s.chequeNumber, s]));
        const s1001 = byChq.get('1001')!;
        expect(s1001.matchedVia).toBe('KEY');
        expect(s1001.fullyCleared).toBe(true);
        expect(s1001.creditLegs).toHaveLength(1);
        expect(s1001.debitLegs).toHaveLength(1);
        expect(s1001.matchedFils).toBe(100000);
        expect(s1001.settledDays).toBe(22); // 2025-03-10 → 2025-04-01

        const s1005 = byChq.get('1005')!;
        expect(s1005.matchedFils).toBe(600000);
        // Nets to zero GL-wise: credit leg −600,000 offsets debit leg +600,000.
    });

    test('key collision pairs one-for-one: 1005 paid, 1006 outstanding, both flagged', async () => {
        const { outcomes, outstanding } = await fixtureRun();
        const byChq = new Map(outcomes.map((o) => [o.chequeNumber, o]));
        expect(byChq.get('1005')!.state).toBe('PAID');
        expect(byChq.get('1005')!.keyCollision).toBe(true);
        expect(byChq.get('1006')!.state).toBe('OUTSTANDING');
        expect(byChq.get('1006')!.keyCollision).toBe(true);

        // The outstanding credit is the SECOND collision row (FIFO by row number).
        const item1006 = outstanding.find((o) => o.cheque?.chequeNumber === '1006')!;
        expect(item1006.direction).toBe('credit');
        expect(item1006.outstandingFils).toBe(600000);
    });

    test('cheque outcomes match the designed table (ops/batch upgrades arrive in T7)', async () => {
        const { outcomes } = await fixtureRun();
        const state = (chq: string) => outcomes.find((o) => o.chequeNumber === chq)!.state;
        expect(state('1001')).toBe('PAID');
        expect(state('1002')).toBe('OUTSTANDING');
        expect(state('1003')).toBe('PRE_WINDOW'); // issued 2024, no in-window issuance credit
        expect(state('1004')).toBe('OUTSTANDING'); // OPS_PAID upgrade lands in Task 7
        expect(state('1005')).toBe('PAID');
        expect(state('1006')).toBe('OUTSTANDING');
        expect(state('1007')).toBe('OUTSTANDING'); // PAID_VIA_BATCH lands in Task 7
        expect(state('1008')).toBe('OUTSTANDING');
        expect(state('1009')).toBe('OUTSTANDING');
        expect(state('1010')).toBe('REGISTER_MATCHED_NO_DEBIT');
        expect(state('1011')).toBe('OUTSTANDING');
    });

    test('issuance-matched credits carry cheque attributes; non-issuance credits do not', async () => {
        const { outstanding } = await fixtureRun();
        const withCheque = outstanding.filter((o) => o.cheque !== undefined);
        // 1002, 1004, 1006, 1007, 1008, 1009, 1010, 1011 — issuance matched, no set formed.
        expect(withCheque).toHaveLength(8);
        const takeOn = outstanding.find((o) => o.journalNumber === '999999999')!;
        expect(takeOn.cheque).toBeUndefined();
        const redeem = outstanding.find((o) => o.journalNumber === '9100')!;
        expect(redeem.cheque).toBeUndefined();
    });

    test('the engine invariant holds: Σ signed outstanding = Σ signed postings', async () => {
        const result = await ingest(readFileSync(FIXTURE), { filename: 'register-sample.xlsx' });
        const { outstanding, summary } = matchRegister(result.postings, result.cheques!, { asOf: AS_OF });
        const signedOutstanding = outstanding.reduce(
            (s, o) => s + (o.direction === 'debit' ? o.outstandingFils : -o.outstandingFils),
            0
        );
        const signedPostings = result.postings.reduce((s, p) => s + p.amountBhdFils, 0);
        expect(signedOutstanding).toBe(signedPostings);
        expect(signedOutstanding).toBe(-2765500);
        expect(summary.netOutstandingFils).toBe(-2765500);
    });

    test('aging keys off the register issuance date for cheques, post date otherwise', async () => {
        const { outstanding } = await fixtureRun();
        const item1011 = outstanding.find((o) => o.cheque?.chequeNumber === '1011')!;
        expect(item1011.ageBucket).toBe('old'); // issued 2025-01-10, 389 days before asOf
        const item1002 = outstanding.find((o) => o.cheque?.chequeNumber === '1002')!;
        expect(item1002.ageBucket).toBe('current');
        const takeOn = outstanding.find((o) => o.journalNumber === '999999999')!;
        expect(takeOn.ageBucket).toBe('old'); // posted 2025-01-01 — post-date aging
    });

    test('summary reports the tuple match key and set counts', async () => {
        const { summary } = await fixtureRun();
        expect(summary.asOf).toBe(AS_OF);
        expect(summary.matchKey).toEqual(['transactionDate', 'journalNumber', 'amountFils']);
        expect(summary.matchedFils).toBe(700000);
        expect(summary.matchedSetCount).toBe(2);
        expect(summary.fullyClearedSetCount).toBe(2);
        expect(summary.outstandingCount).toBe(12); // 10 credits + 2 debits
    });
});

describe('matchRegister — micro-cases', () => {
    test('posting-date variant pass matches when the transaction date drifts', () => {
        const credit = posting({
            amountBhdFils: -50000,
            transactionDate: '2025-06-03', // drifted vs the register's post date
            postDate: '2025-06-05',
            journalNumber: '7777',
        });
        const chq = cheque({
            amountFils: 50000,
            chequeNumber: '9001',
            issuedDate: '2025-06-03',
            issuedPostDate: '2025-06-05', // matches the POSTING date, not the transaction date
            issuedJournal: '7777',
        });
        const { outcomes } = matchRegister([credit], [chq], { asOf: '2026-01-01' });
        expect(outcomes[0].state).toBe('OUTSTANDING'); // issuance found → in population
        expect(outcomes[0].issuanceRowNumber).toBe(credit.rowNumber);
        expect(outcomes[0].matchedVia).toBeUndefined(); // no payment yet
    });

    test('a full variant-resolved pair is flagged POSTING_DATE_VARIANT on the set', () => {
        const credit = posting({
            amountBhdFils: -50000,
            transactionDate: '2025-06-05',
            journalNumber: '7777',
        });
        const debit = posting({
            amountBhdFils: 50000,
            transactionDate: '2025-07-01', // drifted
            postDate: '2025-07-03',
            journalNumber: '8888',
        });
        const chq = cheque({
            amountFils: 50000,
            chequeNumber: '9002',
            issuedPostDate: '2025-06-05',
            issuedJournal: '7777',
            matchedPostDate: '2025-07-03', // matches the debit's POSTING date
            matchedJournal: '8888',
        });
        const { matchedSets, outcomes } = matchRegister([credit, debit], [chq], { asOf: '2026-01-01' });
        expect(matchedSets).toHaveLength(1);
        expect(matchedSets[0].matchedVia).toBe('POSTING_DATE_VARIANT');
        expect(outcomes[0].state).toBe('PAID');
    });

    test('a payment match without an in-window issuance leaves the debit outstanding (PRE_WINDOW)', () => {
        const debit = posting({ amountBhdFils: 80000, transactionDate: '2025-05-01', journalNumber: '6100' });
        const chq = cheque({
            amountFils: 80000,
            chequeNumber: '9003',
            issuedPostDate: '2019-01-01', // issued long before the ledger window
            issuedJournal: '100',
            matchedPostDate: '2025-05-01',
            matchedJournal: '6100',
        });
        const { matchedSets, outstanding, outcomes } = matchRegister([debit], [chq], { asOf: '2026-01-01' });
        expect(matchedSets).toHaveLength(0); // no zero-net set without both legs
        expect(outstanding).toHaveLength(1);
        expect(outstanding[0].direction).toBe('debit');
        expect(outcomes[0].state).toBe('PRE_WINDOW');
        expect(outcomes[0].paymentRowNumbers).toEqual([debit.rowNumber]);
    });

    test('postings after asOf are excluded from the match population', () => {
        const credit = posting({ amountBhdFils: -10000, transactionDate: '2026-06-01', journalNumber: '5' });
        const chq = cheque({ amountFils: 10000, issuedPostDate: '2026-06-01', issuedJournal: '5' });
        const { outstanding, outcomes } = matchRegister([credit], [chq], { asOf: '2026-01-01' });
        expect(outstanding).toHaveLength(0);
        expect(outcomes[0].state).toBe('PRE_WINDOW'); // its issuance is not in the reviewed window
    });
});
