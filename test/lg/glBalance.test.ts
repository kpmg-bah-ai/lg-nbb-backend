import { computeBranchBalances, deriveAsOf, totalBalanceFils } from '../../src/lg/balance';
import { makePosting, signedTotalFils } from './helpers';

describe('deriveAsOf (F3)', () => {
    it('returns the latest post date', () => {
        const postings = [
            makePosting({ amountBhdFils: 1000, postDate: '2024-05-01' }),
            makePosting({ amountBhdFils: 1000, postDate: '2026-06-09' }),
            makePosting({ amountBhdFils: 1000, postDate: '2023-01-08' }),
        ];
        expect(deriveAsOf(postings)).toBe('2026-06-09');
    });

    it('returns undefined for no postings', () => {
        expect(deriveAsOf([])).toBeUndefined();
    });
});

describe('computeBranchBalances (F3)', () => {
    it('sums signed fils per branch, fils-exact', () => {
        const postings = [
            makePosting({ amountBhdFils: 555, branchNumber: '1' }), // 0.555 BHD
            makePosting({ amountBhdFils: 340463, branchNumber: '1' }), // 340.463
            makePosting({ amountBhdFils: -136343, branchNumber: '1' }), // -136.343
            makePosting({ amountBhdFils: 72501861, branchNumber: '2' }),
        ];
        const balances = computeBranchBalances(postings);
        expect(balances).toHaveLength(2);
        const b1 = balances.find((b) => b.branchNumber === '1')!;
        expect(b1.balanceFils).toBe(555 + 340463 - 136343);
        expect(b1.balance).toBe(204.675);
        expect(b1.postingCount).toBe(3);
        const b2 = balances.find((b) => b.branchNumber === '2')!;
        expect(b2.balanceFils).toBe(72501861);
        expect(b2.balance).toBe(72501.861);
    });

    it('excludes postings after the as-at date and tracks first/last post dates', () => {
        const postings = [
            makePosting({ amountBhdFils: 1000, postDate: '2024-01-10' }),
            makePosting({ amountBhdFils: 2000, postDate: '2025-03-05' }),
            makePosting({ amountBhdFils: 4000, postDate: '2026-06-09' }),
        ];
        const balances = computeBranchBalances(postings, '2025-12-31');
        expect(balances).toHaveLength(1);
        expect(balances[0].balanceFils).toBe(3000);
        expect(balances[0].postingCount).toBe(2);
        expect(balances[0].firstPostDate).toBe('2024-01-10');
        expect(balances[0].lastPostDate).toBe('2025-03-05');
    });

    it('includes postings dated exactly on the as-at date', () => {
        const postings = [makePosting({ amountBhdFils: 5000, postDate: '2026-06-30' })];
        expect(computeBranchBalances(postings, '2026-06-30')[0].balanceFils).toBe(5000);
    });

    it('separates entities and GLs, and sorts by gl then branch', () => {
        const postings = [
            makePosting({ amountBhdFils: 100, gl: 'Z9999999', branchNumber: '1' }),
            makePosting({ amountBhdFils: 200, gl: 'D2810085', branchNumber: '2' }),
            makePosting({ amountBhdFils: 300, gl: 'D2810085', branchNumber: '1' }),
        ];
        const balances = computeBranchBalances(postings);
        expect(balances.map((b) => `${b.gl}/${b.branchNumber}`)).toEqual(['D2810085/1', 'D2810085/2', 'Z9999999/1']);
    });

    it('returns an empty list for no postings', () => {
        expect(computeBranchBalances([])).toEqual([]);
    });

    it('totalBalanceFils equals the signed sum of all included postings', () => {
        const postings = [
            makePosting({ amountBhdFils: 5590614, branchNumber: '1' }),
            makePosting({ amountBhdFils: -5590614, branchNumber: '2' }),
            makePosting({ amountBhdFils: 250, branchNumber: '3' }),
        ];
        expect(totalBalanceFils(computeBranchBalances(postings))).toBe(signedTotalFils(postings));
    });
});
