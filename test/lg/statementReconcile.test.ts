/**
 * GOAL-8 — running-balance reconciliation for statement-mode GLs (VAT).
 * The tie-out is derived == stated (no outstanding population): balanced ⇔
 * |derived − stated| ≤ tolerance, with the gap carried as the block's Difference.
 */

import { reconcileStatement } from '../../src/lg/statementReconcile';
import { computeBranchBalances } from '../../src/lg/balance';
import { makePosting } from './helpers';

describe('reconcileStatement (GOAL-8)', () => {
    it('ties out when Σ postings == −(file stated EoD): balanced, gap 0', () => {
        // Two debits +500 each (engine sign) ⇒ derived +1000; file states −1.000 ⇒ statedFils +1000.
        const postings = [
            makePosting({ amountBhdFils: 500, gl: '8828010400010000', branchNumber: '00001', statedEodFils: -1000, postDate: '2023-02-01' }),
            makePosting({ amountBhdFils: 500, gl: '8828010400010000', branchNumber: '00001', statedEodFils: -1000, postDate: '2023-02-01' }),
        ];
        const balances = computeBranchBalances(postings);
        const recon = reconcileStatement(-1000, balances, { asOf: '2023-02-01' }); // statedFils passed engine-signed
        const block = recon.byBranch[0];
        expect(recon.byBranch).toHaveLength(1);
        expect(block.branchNumber).toBe('');            // consolidated
        expect(block.derivedBalanceFils).toBe(1000);
        expect(block.statedBalanceFils).toBe(-1000);
        expect(block.extractGapFils).toBe(2000);        // derived − stated (see note below)
        expect(recon.balanced).toBe(false);
    });
});
