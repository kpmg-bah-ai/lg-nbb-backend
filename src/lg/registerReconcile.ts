/**
 * GOAL-3 R6 — stated-vs-derived balance and the decomposed Difference
 * (GOAL-3 §4.5). Pure functions.
 *
 * The register-family ledger extract STATES the GL balance per row
 * (`End Date EoD Balance`). The engine derives its own balance from the
 * postings and reconciles the two — the reference file itself carries a
 * 34,001.866 BHD gap between derived net and stated EoD, which must surface,
 * never hide (GOAL.md §9.6).
 *
 * The block algebra is the same as breakdown-mode reconcile():
 *     difference = glBalance − Σ signed outstanding
 * with two register-mode twists: glBalance is the STATED balance (falling back
 * to derived when the file states none), and Σ outstanding covers ALL
 * unmatched postings (the invariant population), so the difference IS the
 * reviewer-chase residual after the statement + classified decomposition:
 *     S − O_all = (S − O_stmt) − (O_all − O_stmt) = businessDifference − classified.
 *
 * Sign convention (frozen default, GOAL-3 §11): the file states a credit
 * balance as a positive magnitude (MC PAYABLE is a liability), so the stated
 * figure is negated into engine-signed fils (credit = negative).
 */

import {
    BranchBalance,
    BranchReconciliation,
    filsToBhd,
    ParsedPosting,
    ParseError,
    Reconciliation,
} from '../shared/models';
import { totalBalanceFils } from './balance';
import { DEFAULT_TOLERANCE_FILS } from './reconcile';
import { RegisterMatchResult, statementOutstanding } from './registerMatch';

export interface StatedBalanceResult {
    /** Engine-signed integer fils (credit balances negative); undefined when not stated. */
    statedFils?: number;
    error?: ParseError;
}

/**
 * Reads the stated End Date EoD Balance from the rows dated on the extract's
 * final post date. Those rows must agree — a disagreement is surfaced as
 * INCONSISTENT_STATED_BALANCE and the stated figure is discarded (the caller
 * falls back to the derived balance).
 */
export function extractStatedBalance(postings: ParsedPosting[]): StatedBalanceResult {
    let finalDate: string | undefined;
    for (const p of postings) {
        if (finalDate === undefined || p.postDate > finalDate) {
            finalDate = p.postDate;
        }
    }
    if (finalDate === undefined) {
        return {};
    }
    const stated = new Set<number>();
    for (const p of postings) {
        if (p.postDate === finalDate && p.statedEodFils !== undefined) {
            stated.add(p.statedEodFils);
        }
    }
    if (stated.size === 0) {
        return {};
    }
    if (stated.size > 1) {
        return {
            error: {
                code: 'INCONSISTENT_STATED_BALANCE',
                message:
                    `The final-day rows disagree on the stated End Date EoD Balance ` +
                    `(${[...stated].map((f) => filsToBhd(f).toFixed(3)).join(' vs ')}) — ` +
                    'the stated figure was discarded; reconciling against the derived balance instead',
            },
        };
    }
    // The file states a credit balance as a positive magnitude — negate into
    // engine-signed fils (GOAL-3 §4.5 frozen default, open question §11).
    return { statedFils: -[...stated][0] };
}

export interface RegisterReconcileOptions {
    asOf?: string;
    toleranceFils?: number;
}

/**
 * Builds the consolidated GL-level reconciliation block (branchNumber '' = all
 * branches). Section A/B subtotals come from the STATEMENT population
 * (outstanding cheques only); everything else outstanding is the classified
 * exception mass; the residual difference is what a reviewer must chase.
 */
export function reconcileRegister(
    statedFils: number | undefined,
    balances: BranchBalance[],
    match: RegisterMatchResult,
    options: RegisterReconcileOptions = {}
): Reconciliation {
    const toleranceFils = options.toleranceFils ?? DEFAULT_TOLERANCE_FILS;
    const derivedFils = totalBalanceFils(balances);

    // Σ signed over ALL outstanding (= derived balance, by the tested invariant).
    const outstandingNetFils = match.outstanding.reduce(
        (sum, o) => sum + (o.direction === 'debit' ? o.outstandingFils : -o.outstandingFils),
        0
    );

    // Statement sections: outstanding cheques only (GOAL-3 §4.5).
    const statement = statementOutstanding(match);
    let oldCount = 0;
    let oldFils = 0;
    let currentCount = 0;
    let currentFils = 0;
    let statementNetFils = 0;
    for (const item of statement) {
        statementNetFils += item.direction === 'debit' ? item.outstandingFils : -item.outstandingFils;
        if (item.ageBucket === 'old') {
            oldCount++;
            oldFils += item.outstandingFils;
        } else {
            currentCount++;
            currentFils += item.outstandingFils;
        }
    }

    const entity = balances[0]?.entity ?? match.outstanding[0]?.entity ?? '';
    const gl = balances[0]?.gl ?? match.outstanding[0]?.gl ?? '';
    const glBalanceFils = statedFils ?? derivedFils;
    const differenceFils = glBalanceFils - outstandingNetFils;
    const classifiedFils = outstandingNetFils - statementNetFils;

    const block: BranchReconciliation = {
        entity,
        gl,
        /** '' = the consolidated all-branches block (the file states one EoD per GL). */
        branchNumber: '',
        glBalanceFils,
        outstandingNetFils,
        outstandingCount: statement.length,
        oldCount,
        oldFils,
        currentCount,
        currentFils,
        differenceFils,
        difference: filsToBhd(differenceFils),
        balanced: Math.abs(differenceFils) <= toleranceFils,
        statedBalanceFils: statedFils,
        derivedBalanceFils: derivedFils,
        extractGapFils: statedFils !== undefined ? derivedFils - statedFils : undefined,
        classifiedFils,
        residualFils: differenceFils,
    };

    return {
        asOf: options.asOf ?? match.summary.asOf,
        toleranceFils,
        balanced: block.balanced,
        totalAbsDifferenceFils: Math.abs(differenceFils),
        byBranch: [block],
    };
}
