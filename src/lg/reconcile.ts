/**
 * LG reconciliation — Difference & Balanced status (GOAL.md §4 F5). Pure functions.
 *
 * Per branch: `Difference = GL Balance − Σ(signed outstanding)`, Balanced ⇔
 * |Difference| ≤ tolerance (default 1 fil = 0.001 BHD, GOAL.md §6). When both sides
 * derive from the same postings the engine's invariant makes the Difference zero by
 * construction; it becomes non-zero when the GL balance is supplied externally
 * (§9.6) or outstanding items are adjusted — which is exactly what must be surfaced,
 * never rounded away (GOAL.md §5).
 *
 * Old/current subtotals are Σ|outstanding| per age bucket — the statement's Section
 * A/B subtotals (amounts are displayed as magnitudes, like the reference sample).
 */

import { BranchBalance, BranchReconciliation, filsToBhd, OutstandingItem, Reconciliation } from '../shared/models';

/** GOAL.md §6: exact-to-the-fil by default. */
export const DEFAULT_TOLERANCE_FILS = 1;

export interface ReconcileOptions {
    asOf?: string;
    toleranceFils?: number;
}

/** Builds the per-branch reconciliation blocks from F3 balances and F4 outstanding items. */
export function reconcile(
    balances: BranchBalance[],
    outstanding: OutstandingItem[],
    options: ReconcileOptions = {}
): Reconciliation {
    const toleranceFils = options.toleranceFils ?? DEFAULT_TOLERANCE_FILS;

    interface Acc extends BranchReconciliation {
        key: string;
    }
    const byKey = new Map<string, Acc>();
    const keyOf = (entity: string, gl: string, branchNumber: string) => JSON.stringify([entity, gl, branchNumber]);
    const entryFor = (entity: string, gl: string, branchNumber: string): Acc => {
        const key = keyOf(entity, gl, branchNumber);
        let acc = byKey.get(key);
        if (!acc) {
            acc = {
                key,
                entity,
                gl,
                branchNumber,
                glBalanceFils: 0,
                outstandingNetFils: 0,
                outstandingCount: 0,
                oldCount: 0,
                oldFils: 0,
                currentCount: 0,
                currentFils: 0,
                differenceFils: 0,
                difference: 0,
                balanced: true,
            };
            byKey.set(key, acc);
        }
        return acc;
    };

    for (const balance of balances) {
        entryFor(balance.entity, balance.gl, balance.branchNumber).glBalanceFils = balance.balanceFils;
    }
    for (const item of outstanding) {
        const acc = entryFor(item.entity, item.gl, item.branchNumber);
        acc.outstandingCount++;
        acc.outstandingNetFils += item.direction === 'debit' ? item.outstandingFils : -item.outstandingFils;
        if (item.ageBucket === 'old') {
            acc.oldCount++;
            acc.oldFils += item.outstandingFils;
        } else {
            acc.currentCount++;
            acc.currentFils += item.outstandingFils;
        }
    }

    const byBranch: BranchReconciliation[] = [...byKey.values()]
        .map(({ key, ...acc }) => {
            void key;
            const differenceFils = acc.glBalanceFils - acc.outstandingNetFils;
            return {
                ...acc,
                differenceFils,
                difference: filsToBhd(differenceFils),
                balanced: Math.abs(differenceFils) <= toleranceFils,
            };
        })
        .sort((a, b) => a.gl.localeCompare(b.gl) || a.branchNumber.localeCompare(b.branchNumber));

    return {
        asOf: options.asOf,
        toleranceFils,
        balanced: byBranch.every((b) => b.balanced),
        totalAbsDifferenceFils: byBranch.reduce((sum, b) => sum + Math.abs(b.differenceFils), 0),
        byBranch,
    };
}
