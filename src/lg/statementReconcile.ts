/**
 * GOAL-8 — running-balance reconciliation for statement-mode GLs (VAT).
 *
 * There is no matching and no outstanding population. The reconciliation is the
 * derivedEqualsStated tie-out: the balance the postings derive (Σ signed postings)
 * must equal the balance the ledger states (the negated End-Date EoD). Balanced ⇔
 * |derived − stated| ≤ tolerance. The single consolidated block (branchNumber '')
 * carries the gap as its Difference so the frontend banner shows the tie-out.
 */
import { BranchBalance, BranchReconciliation, filsToBhd, Reconciliation } from '../shared/models';
import { totalBalanceFils } from './balance';
import { DEFAULT_TOLERANCE_FILS } from './reconcile';

export interface StatementReconcileOptions {
    asOf?: string;
    toleranceFils?: number;
}

/** `statedFils` must be engine-signed (as extractStatedBalance returns it); undefined ⇒ tie to derived. */
export function reconcileStatement(
    statedFils: number | undefined,
    balances: BranchBalance[],
    options: StatementReconcileOptions = {}
): Reconciliation {
    const toleranceFils = options.toleranceFils ?? DEFAULT_TOLERANCE_FILS;
    const derivedFils = totalBalanceFils(balances);
    const glBalanceFils = statedFils ?? derivedFils;
    const extractGapFils = statedFils !== undefined ? derivedFils - statedFils : 0;

    const block: BranchReconciliation = {
        entity: balances[0]?.entity ?? '',
        gl: balances[0]?.gl ?? '',
        branchNumber: '', // consolidated: one running-balance account across all branches
        glBalanceFils,
        outstandingNetFils: 0,
        outstandingCount: 0,
        oldCount: 0,
        oldFils: 0,
        currentCount: 0,
        currentFils: 0,
        // Difference IS the extract gap in statement mode (no outstanding items).
        differenceFils: extractGapFils,
        difference: filsToBhd(extractGapFils),
        balanced: Math.abs(extractGapFils) <= toleranceFils,
        statedBalanceFils: statedFils,
        derivedBalanceFils: derivedFils,
        extractGapFils: statedFils !== undefined ? extractGapFils : undefined,
        classifiedFils: 0,
        residualFils: extractGapFils,
    };

    return {
        asOf: options.asOf,
        toleranceFils,
        balanced: block.balanced,
        totalAbsDifferenceFils: Math.abs(extractGapFils),
        byBranch: [block],
    };
}
