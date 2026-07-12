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
        expect(matchedSets).toHaveLength(3); // 1001, 1005 (KEY) + the D3 batch set
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

    test('cheque outcomes match the designed table (GOAL-3 Task 1)', async () => {
        const { outcomes } = await fixtureRun();
        const state = (chq: string) => outcomes.find((o) => o.chequeNumber === chq)!.state;
        expect(state('1001')).toBe('PAID');
        expect(state('1002')).toBe('OUTSTANDING');
        expect(state('1003')).toBe('PRE_WINDOW'); // issued 2024, no in-window issuance credit
        expect(state('1004')).toBe('OPS_PAID'); // reviewer disposition — excluded from statement
        expect(state('1005')).toBe('PAID');
        expect(state('1006')).toBe('OUTSTANDING');
        expect(state('1007')).toBe('PAID_VIA_BATCH'); // resolved from D3's Ref.# list
        expect(state('1008')).toBe('PAID_VIA_BATCH');
        expect(state('1009')).toBe('OUTSTANDING');
        expect(state('1010')).toBe('REGISTER_MATCHED_NO_DEBIT');
        expect(state('1011')).toBe('OUTSTANDING');
    });

    test('the batch debit D3 clears 1007+1008 as one set: 2 credit legs + 1 debit leg, nets zero', async () => {
        const { matchedSets, outcomes } = await fixtureRun();
        expect(matchedSets).toHaveLength(3);
        const batch = matchedSets.find((s) => s.matchedVia === 'BATCH_REF')!;
        expect(batch.creditLegs).toHaveLength(2);
        expect(batch.debitLegs).toHaveLength(1);
        expect(batch.matchedFils).toBe(200000);
        expect(batch.fullyCleared).toBe(true);
        expect(batch.debitLegs[0].journalNumber).toBe('8001');
        const paid1007 = outcomes.find((o) => o.chequeNumber === '1007')!;
        expect(paid1007.matchedVia).toBe('BATCH_REF');
        expect(paid1007.paymentRowNumbers).toEqual([batch.debitLegs[0].rowNumber]);
    });

    test('ops-PAID keeps the GL credit outstanding but the statement helper excludes it', async () => {
        const result = await fixtureRun();
        const item1004 = result.outstanding.find((o) => o.cheque?.chequeNumber === '1004')!;
        expect(item1004.outstandingFils).toBe(500000); // invariant: the credit is NOT netted away

        const { statementOutstanding } = await import('../../src/lg/registerMatch');
        const statement = statementOutstanding(result);
        const statementChqs = statement.map((o) => o.cheque!.chequeNumber).sort();
        expect(statementChqs).toEqual(['1002', '1006', '1009', '1011']);
        const statementFils = statement.reduce((s, o) => s + o.outstandingFils, 0);
        expect(statementFils).toBe(925500); // Task-1 sections: 30.000 old + 895.500 current
    });

    test('issuance-matched credits carry cheque attributes; non-issuance credits do not', async () => {
        const { outstanding } = await fixtureRun();
        const withCheque = outstanding.filter((o) => o.cheque !== undefined);
        // 1002, 1004, 1006, 1009, 1010, 1011 — issuance matched, not cleared (1007/1008 cleared via batch).
        expect(withCheque).toHaveLength(6);
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
        expect(summary.matchedFils).toBe(900000); // 100 + 600 + 200 (batch)
        expect(summary.matchedSetCount).toBe(3);
        expect(summary.fullyClearedSetCount).toBe(3);
        expect(summary.outstandingCount).toBe(9); // 8 credits + 1 debit (D4)
    });
});

describe('parseBatchRefs', () => {
    test('extracts and dedups Ref.# journal lists', async () => {
        const { parseBatchRefs } = await import('../../src/lg/registerMatch');
        expect(parseBatchRefs('DEBIT POSTING-20-Ref.# 5007,Ref.# 5008')).toEqual(['5007', '5008']);
        expect(parseBatchRefs('Ref.#100242402,Ref.# 100242402,Ref # 99')).toEqual(['100242402', '99']);
        expect(parseBatchRefs('DEBIT POSTING-20-MISC')).toEqual([]);
        expect(parseBatchRefs(undefined)).toEqual([]);
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

    test('a partial batch allocation keeps the residual as an outstanding debit fragment', () => {
        const credit = posting({ amountBhdFils: -120000, transactionDate: '2025-06-01', journalNumber: '5100' });
        const batchDebit = posting({
            amountBhdFils: 200000,
            transactionDate: '2025-07-01',
            journalNumber: '8100',
            detailedDescription: 'DEBIT POSTING-20-Ref.# 5100',
        });
        const chq = cheque({
            amountFils: 120000,
            chequeNumber: '9010',
            issuedDate: '2025-06-01',
            issuedPostDate: '2025-06-01',
            issuedJournal: '5100',
        });
        const { matchedSets, outstanding, outcomes } = matchRegister([credit, batchDebit], [chq], {
            asOf: '2026-01-01',
        });
        expect(outcomes[0].state).toBe('PAID_VIA_BATCH');
        expect(matchedSets).toHaveLength(1);
        expect(matchedSets[0].matchedFils).toBe(120000);
        expect(matchedSets[0].fullyCleared).toBe(false); // the debit leg is only partly consumed
        expect(matchedSets[0].debitLegs[0].matchedFils).toBe(120000);
        expect(matchedSets[0].debitLegs[0].originalFils).toBe(200000);

        expect(outstanding).toHaveLength(1);
        const residual = outstanding[0];
        expect(residual.reason).toBe('PARTIALLY_MATCHED_DEBIT');
        expect(residual.originalFils).toBe(200000);
        expect(residual.outstandingFils).toBe(80000);
        expect(residual.batchRefs).toEqual(['5100']);

        // Invariant with a partial: −120,000 + 200,000 = +80,000 = Σ signed outstanding.
        const signed = outstanding.reduce(
            (s, o) => s + (o.direction === 'debit' ? o.outstandingFils : -o.outstandingFils),
            0
        );
        expect(signed).toBe(80000);
    });

    test('batch allocation never clears more than the debit amount', () => {
        const c1 = posting({ amountBhdFils: -150000, transactionDate: '2025-06-01', journalNumber: '5200' });
        const c2 = posting({ amountBhdFils: -100000, transactionDate: '2025-06-02', journalNumber: '5201' });
        const batchDebit = posting({
            amountBhdFils: 200000, // can absorb only ONE of the two candidates
            transactionDate: '2025-07-01',
            journalNumber: '8200',
            detailedDescription: 'Ref.# 5200,Ref.# 5201',
        });
        const chq1 = cheque({
            amountFils: 150000,
            chequeNumber: '9020',
            issuedDate: '2025-06-01',
            issuedPostDate: '2025-06-01',
            issuedJournal: '5200',
        });
        const chq2 = cheque({
            amountFils: 100000,
            chequeNumber: '9021',
            issuedDate: '2025-06-02',
            issuedPostDate: '2025-06-02',
            issuedJournal: '5201',
        });
        const { outcomes, outstanding } = matchRegister([c1, c2, batchDebit], [chq1, chq2], { asOf: '2026-01-01' });
        // FIFO by issue date: 9020 (150k) allocates; 9021 (100k) would overshoot 200k — left outstanding.
        expect(outcomes.find((o) => o.chequeNumber === '9020')!.state).toBe('PAID_VIA_BATCH');
        expect(outcomes.find((o) => o.chequeNumber === '9021')!.state).toBe('OUTSTANDING');
        const signed = outstanding.reduce(
            (s, o) => s + (o.direction === 'debit' ? o.outstandingFils : -o.outstandingFils),
            0
        );
        expect(signed).toBe(-150000 - 100000 + 200000); // invariant preserved
    });

    test('an ops-PAID cheque with ledger evidence still resolves via batch, not ops', () => {
        const credit = posting({ amountBhdFils: -50000, transactionDate: '2025-06-01', journalNumber: '5300' });
        const batchDebit = posting({
            amountBhdFils: 50000,
            transactionDate: '2025-07-01',
            journalNumber: '8300',
            detailedDescription: 'Ref.# 7300', // the ops journal, not the issuance journal
        });
        const chq = cheque({
            amountFils: 50000,
            chequeNumber: '9030',
            issuedDate: '2025-06-01',
            issuedPostDate: '2025-06-01',
            issuedJournal: '5300',
            opsRemark: 'PAID',
            opsPaid: true,
            opsJournal: '7300',
        });
        const { outcomes } = matchRegister([credit, batchDebit], [chq], { asOf: '2026-01-01' });
        expect(outcomes[0].state).toBe('PAID_VIA_BATCH'); // ledger evidence beats the manual note
    });

    test('ops-PAID without ledger evidence becomes OPS_PAID', () => {
        const credit = posting({ amountBhdFils: -50000, transactionDate: '2025-06-01', journalNumber: '5400' });
        const chq = cheque({
            amountFils: 50000,
            chequeNumber: '9040',
            issuedDate: '2025-06-01',
            issuedPostDate: '2025-06-01',
            issuedJournal: '5400',
            opsRemark: 'PAID',
            opsPaid: true,
            opsJournal: '7400',
        });
        const { outcomes, outstanding } = matchRegister([credit], [chq], { asOf: '2026-01-01' });
        expect(outcomes[0].state).toBe('OPS_PAID');
        expect(outstanding).toHaveLength(1); // the credit stays in the engine's outstanding set
        expect(outstanding[0].cheque?.chequeNumber).toBe('9040');
    });
});
