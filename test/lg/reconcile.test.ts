import { computeBranchBalances } from '../../src/lg/balance';
import { matchPostings } from '../../src/lg/match';
import { DEFAULT_TOLERANCE_FILS, reconcile } from '../../src/lg/reconcile';
import { BranchBalance, OutstandingItem } from '../../src/shared/models';
import { makePosting } from './helpers';

const K = 1_000_000;

function balance(overrides: Partial<BranchBalance>): BranchBalance {
    return {
        entity: 'BH',
        gl: 'D2810085',
        branchNumber: '1',
        balanceFils: 0,
        balance: 0,
        postingCount: 0,
        ...overrides,
    };
}

function item(overrides: Partial<OutstandingItem> & { outstandingFils: number }): OutstandingItem {
    return {
        entity: 'BH',
        gl: 'D2810085',
        branchNumber: '1',
        postDate: '2026-01-15',
        direction: 'credit',
        originalFils: overrides.outstandingFils,
        outstanding: overrides.outstandingFils / 1000,
        journalNumber: 'J1',
        rowNumber: 1,
        ageBucket: 'current',
        reason: 'UNMATCHED_CREDIT',
        ...overrides,
    };
}

describe('reconcile (F5) — the balanced case', () => {
    it('is Balanced with zero Difference when balances and outstanding derive from the same postings', () => {
        const postings = [
            makePosting({ amountBhdFils: -11 * K, branchNumber: '1', accountNumber: 'A', postDate: '2024-01-10' }),
            makePosting({ amountBhdFils: 9 * K, branchNumber: '1', accountNumber: 'A', postDate: '2025-02-01' }),
            makePosting({ amountBhdFils: -3 * K, branchNumber: '2', accountNumber: 'B', postDate: '2026-05-01' }),
        ];
        const asOf = '2026-06-30';
        const { outstanding } = matchPostings(postings, { asOf });
        const result = reconcile(computeBranchBalances(postings, asOf), outstanding, { asOf });

        expect(result.balanced).toBe(true);
        expect(result.totalAbsDifferenceFils).toBe(0);
        expect(result.byBranch).toHaveLength(2);
        for (const branch of result.byBranch) {
            expect(branch.differenceFils).toBe(0);
            expect(branch.balanced).toBe(true);
        }
        // Branch 1: 2k credit residual left from the 11k−9k partial clear, > 1 year old.
        const b1 = result.byBranch.find((b) => b.branchNumber === '1')!;
        expect(b1.glBalanceFils).toBe(-2 * K);
        expect(b1.outstandingNetFils).toBe(-2 * K);
        expect(b1.oldCount).toBe(1);
        expect(b1.oldFils).toBe(2 * K);
        expect(b1.currentCount).toBe(0);
    });
});

describe('reconcile (F5) — the not-balanced case (external/adjusted figures)', () => {
    it('surfaces a non-zero Difference exactly, never rounded away', () => {
        const balances = [balance({ balanceFils: -72_501_861 })];
        const items = [item({ outstandingFils: 72_501_361 })]; // 0.500 BHD short
        const result = reconcile(balances, items);
        expect(result.balanced).toBe(false);
        const b = result.byBranch[0];
        expect(b.outstandingNetFils).toBe(-72_501_361);
        expect(b.differenceFils).toBe(-72_501_861 - -72_501_361);
        expect(b.differenceFils).toBe(-500);
        expect(b.difference).toBe(-0.5);
        expect(result.totalAbsDifferenceFils).toBe(500);
    });

    it('a branch with a balance but nothing outstanding is a naked difference', () => {
        const result = reconcile([balance({ balanceFils: 125_000 })], []);
        expect(result.byBranch[0].differenceFils).toBe(125_000);
        expect(result.byBranch[0].balanced).toBe(false);
    });

    it('a branch with outstanding items but no balance entry still reconciles (zero balance)', () => {
        const result = reconcile([], [item({ outstandingFils: 5_000, direction: 'debit', reason: 'UNMATCHED_DEBIT' })]);
        expect(result.byBranch).toHaveLength(1);
        expect(result.byBranch[0].glBalanceFils).toBe(0);
        expect(result.byBranch[0].differenceFils).toBe(-5_000);
        expect(result.balanced).toBe(false);
    });
});

describe('reconcile (F5) — tolerance (GOAL §6: 0.001 BHD)', () => {
    it('treats |difference| of exactly one fil as Balanced, two fils as not', () => {
        const oneFil = reconcile([balance({ balanceFils: -1 })], []);
        expect(oneFil.byBranch[0].differenceFils).toBe(-1);
        expect(oneFil.byBranch[0].balanced).toBe(true);

        const twoFils = reconcile([balance({ balanceFils: 2 })], []);
        expect(twoFils.byBranch[0].balanced).toBe(false);
        expect(DEFAULT_TOLERANCE_FILS).toBe(1);
    });

    it('honours a custom tolerance', () => {
        const result = reconcile([balance({ balanceFils: 900 })], [], { toleranceFils: 1000 });
        expect(result.byBranch[0].balanced).toBe(true);
    });
});

describe('reconcile (F5) — section subtotals', () => {
    it('splits old vs current subtotals as magnitudes (statement Sections A/B)', () => {
        const items = [
            item({ outstandingFils: 27_753_361, ageBucket: 'old' }),
            item({ outstandingFils: 44_748_500, ageBucket: 'current', rowNumber: 2 }),
        ];
        const result = reconcile([balance({ balanceFils: -72_501_861 })], items);
        const b = result.byBranch[0];
        expect(b.oldFils).toBe(27_753_361);
        expect(b.currentFils).toBe(44_748_500);
        expect(b.oldFils + b.currentFils).toBe(72_501_861);
        expect(b.differenceFils).toBe(0);
        expect(b.balanced).toBe(true);
    });
});
