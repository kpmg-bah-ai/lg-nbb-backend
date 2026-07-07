import { computeBranchBalances, totalBalanceFils } from '../../src/lg/balance';
import { DEFAULT_MATCH_KEY, matchPostings } from '../../src/lg/match';
import { makePosting, signedTotalFils } from './helpers';

const K = 1_000_000; // 1k BHD in fils

describe('matchPostings (F4) — cleared sets', () => {
    it('clears an exact 1:1 debit/credit pair', () => {
        const postings = [
            makePosting({ amountBhdFils: -5 * K, postDate: '2025-01-10' }),
            makePosting({ amountBhdFils: 5 * K, postDate: '2025-02-01' }),
        ];
        const { outstanding, summary } = matchPostings(postings);
        expect(outstanding).toHaveLength(0);
        expect(summary.matchedFils).toBe(5 * K);
        expect(summary.netOutstandingFils).toBe(0);
    });

    it('clears an 11k credit against 2k + 9k debits (one-to-many)', () => {
        const postings = [
            makePosting({ amountBhdFils: -11 * K, postDate: '2025-01-10' }),
            makePosting({ amountBhdFils: 2 * K, postDate: '2025-03-01' }),
            makePosting({ amountBhdFils: 9 * K, postDate: '2025-04-15' }),
        ];
        const { outstanding, summary } = matchPostings(postings);
        expect(outstanding).toHaveLength(0);
        expect(summary.matchedFils).toBe(11 * K);
        expect(summary.outstandingCount).toBe(0);
    });

    it('clears many-to-many combinations that net to zero', () => {
        const postings = [
            makePosting({ amountBhdFils: -7 * K }),
            makePosting({ amountBhdFils: -4 * K }),
            makePosting({ amountBhdFils: 5 * K }),
            makePosting({ amountBhdFils: 6 * K }),
        ];
        const { outstanding, summary } = matchPostings(postings);
        expect(outstanding).toHaveLength(0);
        expect(summary.matchedFils).toBe(11 * K);
    });
});

describe('matchPostings (F4) — outstanding fragments', () => {
    it('reports the 2k residual when an 11k credit is only offset by a 9k debit', () => {
        const credit = makePosting({ amountBhdFils: -11 * K, postDate: '2025-01-10' });
        const debit = makePosting({ amountBhdFils: 9 * K, postDate: '2025-04-15' });
        const { outstanding, summary } = matchPostings([credit, debit]);
        expect(outstanding).toHaveLength(1);
        const fragment = outstanding[0];
        expect(fragment.direction).toBe('credit');
        expect(fragment.reason).toBe('PARTIALLY_MATCHED_CREDIT');
        expect(fragment.outstandingFils).toBe(2 * K);
        expect(fragment.originalFils).toBe(11 * K);
        expect(fragment.rowNumber).toBe(credit.rowNumber);
        expect(summary.matchedFils).toBe(9 * K);
        expect(summary.netOutstandingFils).toBe(-2 * K);
    });

    it('reports a lone debit as UNMATCHED_DEBIT in full', () => {
        const debit = makePosting({ amountBhdFils: 125_000, postDate: '2024-06-15' });
        const { outstanding } = matchPostings([debit]);
        expect(outstanding).toHaveLength(1);
        expect(outstanding[0].reason).toBe('UNMATCHED_DEBIT');
        expect(outstanding[0].outstandingFils).toBe(125_000);
        expect(outstanding[0].originalFils).toBe(125_000);
    });

    it('never nets across different match keys', () => {
        const postings = [
            makePosting({ amountBhdFils: -5 * K, accountNumber: 'ACC-A' }),
            makePosting({ amountBhdFils: 5 * K, accountNumber: 'ACC-B' }),
        ];
        const { outstanding, summary } = matchPostings(postings);
        expect(outstanding).toHaveLength(2);
        expect(summary.matchedFils).toBe(0);
        expect(outstanding.map((o) => o.reason).sort()).toEqual(['UNMATCHED_CREDIT', 'UNMATCHED_DEBIT']);
        // ...but the signed net still ties out to zero (the F5 identity).
        expect(summary.netOutstandingFils).toBe(0);
    });

    it('a custom match key can widen the grouping', () => {
        const postings = [
            makePosting({ amountBhdFils: -5 * K, accountNumber: 'ACC-A' }),
            makePosting({ amountBhdFils: 5 * K, accountNumber: 'ACC-B' }),
        ];
        const { outstanding } = matchPostings(postings, { matchKey: ['entity', 'gl', 'branchNumber'] });
        expect(outstanding).toHaveLength(0);
    });

    it('ignores zero-amount legs (direction from log code) entirely', () => {
        const postings = [makePosting({ amountBhdFils: 0, direction: 'debit' })];
        const { outstanding, summary } = matchPostings(postings);
        expect(outstanding).toHaveLength(0);
        expect(summary.matchedFils).toBe(0);
    });
});

describe('matchPostings (F4) — review-date population (asOf)', () => {
    it('excludes postings after asOf, so a future credit cannot clear a past debit', () => {
        const debit = makePosting({ amountBhdFils: 5000, postDate: '2025-01-01' });
        const futureCredit = makePosting({ amountBhdFils: -5000, postDate: '2026-01-01' });
        const { outstanding, summary } = matchPostings([debit, futureCredit], { asOf: '2025-06-30' });
        expect(outstanding).toHaveLength(1);
        expect(outstanding[0].reason).toBe('UNMATCHED_DEBIT');
        expect(outstanding[0].outstandingFils).toBe(5000);
        expect(summary.matchedFils).toBe(0);
        // The tie-out identity must hold against the F3 balance at the SAME review date.
        expect(summary.netOutstandingFils).toBe(
            totalBalanceFils(computeBranchBalances([debit, futureCredit], '2025-06-30'))
        );
    });

    it('never reports a posting dated after asOf as outstanding', () => {
        const { outstanding, summary } = matchPostings(
            [makePosting({ amountBhdFils: 700, postDate: '2026-05-01' })],
            { asOf: '2026-01-31' }
        );
        expect(outstanding).toHaveLength(0);
        expect(summary.netOutstandingFils).toBe(0);
    });
});

describe('matchPostings (F4) — aging (Old Items vs current)', () => {
    const asOf = '2026-06-30';

    it('flags fragments older than one year as old', () => {
        const { outstanding } = matchPostings([makePosting({ amountBhdFils: 500, postDate: '2024-01-10' })], { asOf });
        expect(outstanding[0].ageBucket).toBe('old');
    });

    it('keeps fragments within a year as current, including the exact 365-day boundary', () => {
        const { outstanding } = matchPostings(
            [
                makePosting({ amountBhdFils: 500, postDate: '2026-06-01', accountNumber: 'ACC-X' }),
                makePosting({ amountBhdFils: 700, postDate: '2025-06-30', accountNumber: 'ACC-Y' }),
            ],
            { asOf }
        );
        expect(outstanding.map((o) => o.ageBucket)).toEqual(['current', 'current']);
    });

    it('treats an exact one-year span across a leap February as old (366 days > 365)', () => {
        // §9.4: "old" is > 365 days; a calendar year containing 29 Feb spans 366.
        const { outstanding } = matchPostings([makePosting({ amountBhdFils: 500, postDate: '2023-06-30' })], {
            asOf: '2024-06-30',
        });
        expect(outstanding[0].ageBucket).toBe('old');
    });

    it('counts old vs current in the summary', () => {
        const { summary } = matchPostings(
            [
                makePosting({ amountBhdFils: 500, postDate: '2022-03-15', accountNumber: 'ACC-X' }),
                makePosting({ amountBhdFils: 700, postDate: '2026-04-08', accountNumber: 'ACC-Y' }),
            ],
            { asOf }
        );
        expect(summary.oldCount).toBe(1);
        expect(summary.currentCount).toBe(1);
    });
});

describe('matchPostings (F4) — the F5 tie-out identity and rollups', () => {
    it('signed outstanding always equals the signed postings total and the F3 balance', () => {
        const postings = [
            makePosting({ amountBhdFils: -11 * K, branchNumber: '1', accountNumber: 'A' }),
            makePosting({ amountBhdFils: 2 * K, branchNumber: '1', accountNumber: 'A' }),
            makePosting({ amountBhdFils: 9 * K, branchNumber: '1', accountNumber: 'A' }),
            makePosting({ amountBhdFils: -3 * K, branchNumber: '1', accountNumber: 'B' }),
            makePosting({ amountBhdFils: 340_463, branchNumber: '2', accountNumber: 'C' }),
            makePosting({ amountBhdFils: -78_500, branchNumber: '2', accountNumber: 'D' }),
        ];
        const { summary } = matchPostings(postings);
        const net = signedTotalFils(postings);
        expect(summary.netOutstandingFils).toBe(net);
        expect(totalBalanceFils(computeBranchBalances(postings))).toBe(net);
    });

    it('rolls outstanding up per branch with signed fils', () => {
        const { summary } = matchPostings([
            makePosting({ amountBhdFils: -3 * K, branchNumber: '1', accountNumber: 'B' }),
            makePosting({ amountBhdFils: 340_463, branchNumber: '2', accountNumber: 'C' }),
            makePosting({ amountBhdFils: -78_500, branchNumber: '2', accountNumber: 'D' }),
        ]);
        expect(summary.byBranch).toEqual([
            { branchNumber: '1', outstandingCount: 1, outstandingFils: -3 * K, matchedSetCount: 0 },
            { branchNumber: '2', outstandingCount: 2, outstandingFils: 340_463 - 78_500, matchedSetCount: 0 },
        ]);
        expect(summary.outstandingDebitFils).toBe(340_463);
        expect(summary.outstandingCreditFils).toBe(3 * K + 78_500);
    });

    it('exposes the match key it used', () => {
        const { summary } = matchPostings([]);
        expect(summary.matchKey).toEqual(DEFAULT_MATCH_KEY.map(String));
    });
});
