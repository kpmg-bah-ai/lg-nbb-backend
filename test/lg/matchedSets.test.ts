import { matchPostings } from '../../src/lg/match';
import { makePosting } from './helpers';

const K = 1_000_000; // 1k BHD in fils

describe('matchPostings (G1) — cleared-set capture', () => {
    it('emits a 1:1 set with both legs and their offsets', () => {
        const credit = makePosting({ amountBhdFils: -5 * K, postDate: '2025-01-10', journalNumber: 'J-CR' });
        const debit = makePosting({ amountBhdFils: 5 * K, postDate: '2025-02-01', journalNumber: 'J-DR' });
        const { matchedSets, summary } = matchPostings([credit, debit]);

        expect(matchedSets).toHaveLength(1);
        const set = matchedSets[0];
        expect(set.matchedFils).toBe(5 * K);
        expect(set.fullyCleared).toBe(true);
        expect(set.creditLegs).toHaveLength(1);
        expect(set.debitLegs).toHaveLength(1);
        expect(set.creditLegs[0]).toMatchObject({
            journalNumber: 'J-CR',
            rowNumber: credit.rowNumber,
            originalFils: 5 * K,
            matchedFils: 5 * K,
            direction: 'credit',
        });
        expect(set.debitLegs[0]).toMatchObject({ journalNumber: 'J-DR', rowNumber: debit.rowNumber });
        expect(summary.matchedSetCount).toBe(1);
        expect(summary.fullyClearedSetCount).toBe(1);
    });

    it('captures a 1:N clear (11k credit = 2k + 9k debits) as one set with settledDays', () => {
        const postings = [
            makePosting({ amountBhdFils: -11 * K, postDate: '2025-01-10' }),
            makePosting({ amountBhdFils: 2 * K, postDate: '2025-03-01' }),
            makePosting({ amountBhdFils: 9 * K, postDate: '2025-04-15' }),
        ];
        const { matchedSets } = matchPostings(postings);

        expect(matchedSets).toHaveLength(1);
        const set = matchedSets[0];
        expect(set.creditLegs).toHaveLength(1);
        expect(set.debitLegs).toHaveLength(2);
        expect(set.matchedFils).toBe(11 * K);
        expect(set.debitLegs.map((l) => l.matchedFils)).toEqual([2 * K, 9 * K]);
        expect(set.firstCreditDate).toBe('2025-01-10');
        expect(set.finalDebitDate).toBe('2025-04-15');
        expect(set.settledDays).toBe(95); // 10 Jan → 15 Apr 2025
        expect(set.fullyCleared).toBe(true);
    });

    it('captures an M:N chain (7k + 4k credits vs 5k + 6k debits) as one set', () => {
        const postings = [
            makePosting({ amountBhdFils: -7 * K, postDate: '2025-01-01' }),
            makePosting({ amountBhdFils: -4 * K, postDate: '2025-01-05' }),
            makePosting({ amountBhdFils: 5 * K, postDate: '2025-02-01' }),
            makePosting({ amountBhdFils: 6 * K, postDate: '2025-02-10' }),
        ];
        const { matchedSets } = matchPostings(postings);

        expect(matchedSets).toHaveLength(1);
        const set = matchedSets[0];
        expect(set.creditLegs).toHaveLength(2);
        expect(set.debitLegs).toHaveLength(2);
        expect(set.matchedFils).toBe(11 * K);
        expect(set.fullyCleared).toBe(true);
        // Per-leg offsets must sum to each leg's original amount when fully cleared.
        for (const leg of [...set.creditLegs, ...set.debitLegs]) {
            expect(leg.matchedFils).toBe(leg.originalFils);
        }
    });

    it('separates independent pairs into distinct sets (equal-remainder boundary)', () => {
        const postings = [
            makePosting({ amountBhdFils: -5 * K, postDate: '2025-01-01' }),
            makePosting({ amountBhdFils: 5 * K, postDate: '2025-01-15' }),
            makePosting({ amountBhdFils: -3 * K, postDate: '2025-02-01' }),
            makePosting({ amountBhdFils: 3 * K, postDate: '2025-02-20' }),
        ];
        const { matchedSets } = matchPostings(postings);

        expect(matchedSets).toHaveLength(2);
        expect(matchedSets.map((s) => s.matchedFils)).toEqual([5 * K, 3 * K]);
        expect(matchedSets.every((s) => s.fullyCleared)).toBe(true);
    });

    it('flags a partial clear (11k credit vs 9k debit) as a not-fully-cleared set plus a residual', () => {
        const credit = makePosting({ amountBhdFils: -11 * K, postDate: '2025-01-10' });
        const debit = makePosting({ amountBhdFils: 9 * K, postDate: '2025-04-15' });
        const { matchedSets, outstanding, summary } = matchPostings([credit, debit]);

        expect(matchedSets).toHaveLength(1);
        const set = matchedSets[0];
        expect(set.fullyCleared).toBe(false);
        expect(set.matchedFils).toBe(9 * K);
        expect(set.creditLegs[0].originalFils).toBe(11 * K);
        expect(set.creditLegs[0].matchedFils).toBe(9 * K);
        // The 2k residual is outstanding — never netted silently into the set.
        expect(outstanding).toHaveLength(1);
        expect(outstanding[0].outstandingFils).toBe(2 * K);
        expect(summary.fullyClearedSetCount).toBe(0);
        expect(summary.matchedSetCount).toBe(1);
    });

    it('Σ set.matchedFils equals summary.matchedFils across mixed scenarios', () => {
        const postings = [
            // group A: full 1:N clear
            makePosting({ amountBhdFils: -11 * K, accountNumber: 'A', postDate: '2025-01-01' }),
            makePosting({ amountBhdFils: 2 * K, accountNumber: 'A', postDate: '2025-01-10' }),
            makePosting({ amountBhdFils: 9 * K, accountNumber: 'A', postDate: '2025-01-20' }),
            // group B: partial
            makePosting({ amountBhdFils: -10 * K, accountNumber: 'B', postDate: '2025-02-01' }),
            makePosting({ amountBhdFils: 4 * K, accountNumber: 'B', postDate: '2025-02-11' }),
            // group C: unmatched only — no set
            makePosting({ amountBhdFils: 700, accountNumber: 'C', postDate: '2025-03-01' }),
        ];
        const { matchedSets, summary } = matchPostings(postings);

        const setTotal = matchedSets.reduce((s, m) => s + m.matchedFils, 0);
        expect(setTotal).toBe(summary.matchedFils);
        expect(setTotal).toBe(11 * K + 4 * K);
        expect(summary.matchedSetCount).toBe(2);
    });

    it('never builds a set across different match keys', () => {
        const postings = [
            makePosting({ amountBhdFils: -5 * K, accountNumber: 'ACC-A' }),
            makePosting({ amountBhdFils: 5 * K, accountNumber: 'ACC-B' }),
        ];
        const { matchedSets } = matchPostings(postings);
        expect(matchedSets).toHaveLength(0);
    });

    it('respects the review-date population: a post-asOf debit cannot appear in a set', () => {
        const credit = makePosting({ amountBhdFils: -5 * K, postDate: '2025-01-01' });
        const futureDebit = makePosting({ amountBhdFils: 5 * K, postDate: '2026-01-01' });
        const { matchedSets, outstanding } = matchPostings([credit, futureDebit], { asOf: '2025-06-30' });

        expect(matchedSets).toHaveLength(0);
        expect(outstanding).toHaveLength(1);
        expect(outstanding[0].reason).toBe('UNMATCHED_CREDIT');
    });

    it('carries the set-level account only when unambiguous and rolls sets up per branch', () => {
        const postings = [
            makePosting({ amountBhdFils: -5 * K, branchNumber: '1', accountNumber: 'ACC-1', postDate: '2025-01-01' }),
            makePosting({ amountBhdFils: 5 * K, branchNumber: '1', accountNumber: 'ACC-1', postDate: '2025-01-05' }),
            makePosting({ amountBhdFils: 900, branchNumber: '2', accountNumber: 'ACC-2', postDate: '2025-01-01' }),
        ];
        const { matchedSets, summary } = matchPostings(postings);

        expect(matchedSets).toHaveLength(1);
        expect(matchedSets[0].accountNumber).toBe('ACC-1');
        expect(matchedSets[0].branchNumber).toBe('1');
        expect(summary.byBranch).toEqual([
            { branchNumber: '1', outstandingCount: 0, outstandingFils: 0, matchedSetCount: 1 },
            { branchNumber: '2', outstandingCount: 1, outstandingFils: 900, matchedSetCount: 0 },
        ]);
    });
});
