/**
 * GOAL-3 R6 — stated-vs-derived balance and the decomposed Difference
 * (src/lg/registerReconcile.ts).
 *
 * Signed-fils algebra on the fixture (GOAL-3 §7 Task 1, engine convention:
 * credits negative):
 *   stated S        = −2,730,000   (file states 2,730.000 credit)
 *   derived / Σout  = −2,765,500   (invariant: Σ signed outstanding = derived)
 *   statement O_stmt= −925,500     (cheques 1002/1006/1009/1011)
 *   classified C    = O_all − O_stmt = −1,840,000
 *   difference      = S − O_all = +35,500  (= residual, the reviewer-chase number)
 *   extract gap     = derived − stated = −35,500
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { computeBranchBalances } from '../../src/lg/balance';
import { ingest } from '../../src/lg/ingest';
import { matchRegister } from '../../src/lg/registerMatch';
import { extractStatedBalance, reconcileRegister } from '../../src/lg/registerReconcile';
import { ParsedPosting } from '../../src/shared/models';

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

async function fixtureParts() {
    const result = await ingest(readFileSync(FIXTURE), { filename: 'register-sample.xlsx' });
    const match = matchRegister(result.postings, result.cheques!, { asOf: AS_OF });
    const balances = computeBranchBalances(result.postings, AS_OF);
    const stated = extractStatedBalance(result.postings);
    return { result, match, balances, stated };
}

describe('extractStatedBalance', () => {
    test('reads the End Date EoD Balance from final-day rows, engine-signed', async () => {
        const { stated } = await fixtureParts();
        expect(stated.error).toBeUndefined();
        expect(stated.statedFils).toBe(-2730000); // file's 2,730.000 credit magnitude, negated
    });

    test('disagreeing final-day rows are INCONSISTENT_STATED_BALANCE', () => {
        const rows = [
            posting({ amountBhdFils: -100000, postDate: '2025-12-31', statedEodFils: 2730000 }),
            posting({ amountBhdFils: 50000, postDate: '2025-12-31', statedEodFils: 999000 }),
        ];
        const { statedFils, error } = extractStatedBalance(rows);
        expect(statedFils).toBeUndefined();
        expect(error).toEqual(expect.objectContaining({ code: 'INCONSISTENT_STATED_BALANCE' }));
    });

    test('no stated values at all is simply undefined, not an error', () => {
        const rows = [posting({ amountBhdFils: -100000, postDate: '2025-12-31' })];
        const { statedFils, error } = extractStatedBalance(rows);
        expect(statedFils).toBeUndefined();
        expect(error).toBeUndefined();
    });
});

describe('reconcileRegister — fixture decomposition (GOAL-3 Task 1 table)', () => {
    test('produces one GL-level block with the exact decomposition', async () => {
        const { match, balances, stated } = await fixtureParts();
        const reconciliation = reconcileRegister(stated.statedFils, balances, match, { asOf: AS_OF });

        expect(reconciliation.asOf).toBe(AS_OF);
        expect(reconciliation.toleranceFils).toBe(1);
        expect(reconciliation.byBranch).toHaveLength(1); // consolidated GL-level block

        const block = reconciliation.byBranch[0];
        expect(block.gl).toBe('99801000');
        expect(block.branchNumber).toBe(''); // '' = all branches (consolidated)

        expect(block.statedBalanceFils).toBe(-2730000);
        expect(block.derivedBalanceFils).toBe(-2765500);
        expect(block.extractGapFils).toBe(-35500); // derived − stated

        expect(block.glBalanceFils).toBe(-2730000); // stated drives the block
        expect(block.outstandingNetFils).toBe(-2765500); // Σ signed ALL outstanding
        expect(block.differenceFils).toBe(35500);
        expect(block.residualFils).toBe(35500); // the reviewer-chase number
        expect(block.classifiedFils).toBe(-1840000); // ops-paid + non-issuance + reg-paid-no-debit − stray debit

        // Statement sections (magnitudes): CHQ 1011 old, 1002/1006/1009 current.
        expect(block.oldCount).toBe(1);
        expect(block.oldFils).toBe(30000);
        expect(block.currentCount).toBe(3);
        expect(block.currentFils).toBe(895500);
        expect(block.outstandingCount).toBe(4);

        expect(block.balanced).toBe(false); // |35,500| > 1 fil
        expect(reconciliation.balanced).toBe(false);
        expect(reconciliation.totalAbsDifferenceFils).toBe(35500);
    });

    test('classified + statement decomposition ties back to the derived balance', async () => {
        const { match, balances, stated } = await fixtureParts();
        const block = reconcileRegister(stated.statedFils, balances, match, { asOf: AS_OF }).byBranch[0];
        // O_all = O_stmt + C  ⇒  −2,765,500 = −925,500 + (−1,840,000)
        const statementNet = -(block.oldFils + block.currentFils);
        expect(statementNet + block.classifiedFils!).toBe(block.outstandingNetFils);
    });

    test('without a stated balance the block falls back to derived and balances', async () => {
        const { match, balances } = await fixtureParts();
        const reconciliation = reconcileRegister(undefined, balances, match, { asOf: AS_OF });
        const block = reconciliation.byBranch[0];
        expect(block.statedBalanceFils).toBeUndefined();
        expect(block.glBalanceFils).toBe(-2765500); // derived
        expect(block.differenceFils).toBe(0); // invariant makes this exact
        expect(block.extractGapFils).toBeUndefined();
        expect(block.balanced).toBe(true);
    });

    test('the exact tolerance boundary: one fil of residual still balances', async () => {
        const { match, balances } = await fixtureParts();
        // Derived is −2,765,500; a stated balance one fil away must still balance.
        const reconciliation = reconcileRegister(-2765501, balances, match, { asOf: AS_OF });
        expect(reconciliation.byBranch[0].differenceFils).toBe(-1);
        expect(reconciliation.byBranch[0].balanced).toBe(true);
        expect(reconciliation.balanced).toBe(true);
    });
});
