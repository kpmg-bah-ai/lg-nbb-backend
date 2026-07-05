/**
 * LG reconciliation — GL balance computation (GOAL.md §4 F3). Pure functions.
 *
 * The balance of a clearing/suspense account per (entity, gl, branch) is the sum of
 * its signed posting amounts up to the review date. All arithmetic is integer fils
 * (GOAL.md §6); decimals exist for display only.
 */

import { BranchBalance, filsToBhd, ParsedPosting } from '../shared/models';

/** The latest post date across the postings — the natural default review date. */
export function deriveAsOf(postings: ParsedPosting[]): string | undefined {
    let max: string | undefined;
    for (const p of postings) {
        if (max === undefined || p.postDate > max) {
            max = p.postDate;
        }
    }
    return max;
}

/**
 * Per-(entity, gl, branch) balances as at `asOf` (inclusive; postings after it are
 * excluded). ISO yyyy-mm-dd strings compare correctly lexicographically.
 */
export function computeBranchBalances(postings: ParsedPosting[], asOf?: string): BranchBalance[] {
    interface Acc {
        entity: string;
        gl: string;
        branchNumber: string;
        balanceFils: number;
        postingCount: number;
        firstPostDate?: string;
        lastPostDate?: string;
    }
    const groups = new Map<string, Acc>();
    for (const p of postings) {
        if (asOf !== undefined && p.postDate > asOf) {
            continue;
        }
        const key = JSON.stringify([p.entity, p.gl, p.branchNumber]);
        let acc = groups.get(key);
        if (!acc) {
            acc = { entity: p.entity, gl: p.gl, branchNumber: p.branchNumber, balanceFils: 0, postingCount: 0 };
            groups.set(key, acc);
        }
        acc.balanceFils += p.amountBhdFils;
        acc.postingCount++;
        if (acc.firstPostDate === undefined || p.postDate < acc.firstPostDate) {
            acc.firstPostDate = p.postDate;
        }
        if (acc.lastPostDate === undefined || p.postDate > acc.lastPostDate) {
            acc.lastPostDate = p.postDate;
        }
    }
    return [...groups.values()]
        .map((acc) => ({ ...acc, balance: filsToBhd(acc.balanceFils) }))
        .sort((a, b) => a.gl.localeCompare(b.gl) || a.branchNumber.localeCompare(b.branchNumber));
}

/** Account-level total across branch balances, integer fils. */
export function totalBalanceFils(balances: BranchBalance[]): number {
    return balances.reduce((sum, b) => sum + b.balanceFils, 0);
}
