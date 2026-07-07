/**
 * LG reconciliation — exception classification (GOAL-2 G2 / GOAL.md §3 F6). Pure functions.
 *
 * Every outstanding item becomes exactly ONE exception (exceptions ⊇ outstanding —
 * a debit with no matching credit is always surfaced, never dropped). On top of the
 * base UNMATCHED_* / PARTIALLY_MATCHED_* reasons, two patterns are classified:
 *
 *   DUPLICATE        — two (or more) outstanding postings share journal number, post
 *                      date, amount, account and direction: the "system retry"
 *                      signature (GOAL-2 §8.5). All twins are flagged, linked via
 *                      relatedRowNumbers.
 *   AMOUNT_MISMATCH  — within one branch scope (entity, gl, branchNumber), exactly
 *                      one whole unmatched debit and one whole unmatched credit
 *                      remain and their amounts differ by ≤ the mismatch band
 *                      (default 1.000 BHD): a probable keying error, both sides
 *                      flagged and cross-linked (GOAL-2 §8.2 default proposal: flag
 *                      only when the pair is unambiguous). The scope is the branch,
 *                      not the match-key group: FIFO offsetting exhausts one
 *                      direction per group, so a surviving debit+credit pair is by
 *                      construction in different groups (e.g. mis-keyed accounts).
 *
 * Precedence: DUPLICATE > AMOUNT_MISMATCH > base reason. Classification only ever
 * upgrades the reason/message — amounts, rows and aging pass through untouched, so
 * Σ outstanding fils is preserved exactly (GOAL.md §5: nothing rounded away).
 */

import { ExceptionSummary, filsToBhd, LgException, LgExceptionReason, OutstandingItem } from '../shared/models';

/** GOAL-2 §8.2: near-miss band for AMOUNT_MISMATCH — 1.000 BHD by default. */
export const DEFAULT_MISMATCH_BAND_FILS = 1000;

export interface ExceptionOptions {
    /** |debit − credit| ≤ band ⇒ AMOUNT_MISMATCH candidate (integer fils). */
    mismatchBandFils?: number;
}

export interface ExceptionResult {
    exceptions: LgException[];
    summary: ExceptionSummary;
}

const fmtBhd = (fils: number): string => filsToBhd(fils).toFixed(3);

function baseMessage(item: OutstandingItem): string {
    const opposite = item.direction === 'debit' ? 'credit' : 'debit';
    if (item.reason.startsWith('PARTIALLY')) {
        return (
            `Only ${fmtBhd(item.originalFils - item.outstandingFils)} BHD of this ${item.direction} was offset; ` +
            `${fmtBhd(item.outstandingFils)} BHD remains outstanding as at the review date. ` +
            `Verify the remaining ${opposite} leg(s) or confirm a partial settlement.`
        );
    }
    return (
        `This ${item.direction} posting has no matching ${opposite} in the account as at the review date. ` +
        `It keeps the branch reconciliation from clearing until it is matched or investigated.`
    );
}

function duplicateMessage(item: OutstandingItem, related: number[]): string {
    return (
        `${related.length + 1} ${item.direction} postings share journal ${item.journalNumber}, ` +
        `post date ${item.postDate} and amount ${fmtBhd(item.originalFils)} BHD — a probable system retry. ` +
        `Confirm and reverse the duplicate entry (rows ${[item.rowNumber, ...related].sort((a, b) => a - b).join(', ')}).`
    );
}

function mismatchMessage(item: OutstandingItem, other: OutstandingItem): string {
    return (
        `This ${item.direction} of ${fmtBhd(item.outstandingFils)} BHD nearly matches the opposite ` +
        `${other.direction} of ${fmtBhd(other.outstandingFils)} BHD (row ${other.rowNumber}) in the same branch — ` +
        `difference ${fmtBhd(Math.abs(item.outstandingFils - other.outstandingFils))} BHD. ` +
        `Likely a keying error at the source system; request a correction memo referencing both journals.`
    );
}

/** Classifies the outstanding items into typed, reviewer-readable exceptions. */
export function detectExceptions(outstanding: OutstandingItem[], options: ExceptionOptions = {}): ExceptionResult {
    const band = options.mismatchBandFils ?? DEFAULT_MISMATCH_BAND_FILS;

    const exceptions: LgException[] = outstanding.map((item) => ({
        entity: item.entity,
        gl: item.gl,
        branchNumber: item.branchNumber,
        accountNumber: item.accountNumber,
        postDate: item.postDate,
        direction: item.direction,
        originalFils: item.originalFils,
        outstandingFils: item.outstandingFils,
        logCode: item.logCode,
        journalNumber: item.journalNumber,
        sequence: item.sequence,
        rowNumber: item.rowNumber,
        sheet: item.sheet,
        ageBucket: item.ageBucket,
        reason: item.reason as LgExceptionReason,
        message: baseMessage(item),
    }));

    // DUPLICATE: the system-retry signature. Only whole (un-offset) postings qualify —
    // a partially matched leg has already consumed a counter-leg, so it is not a retry twin.
    const bySignature = new Map<string, number[]>();
    exceptions.forEach((exc, i) => {
        if (outstanding[i].reason.startsWith('PARTIALLY')) {
            return;
        }
        const signature = JSON.stringify([
            exc.entity,
            exc.gl,
            exc.branchNumber,
            exc.accountNumber ?? '',
            exc.journalNumber,
            exc.postDate,
            exc.direction,
            exc.originalFils,
        ]);
        const group = bySignature.get(signature);
        if (group) {
            group.push(i);
        } else {
            bySignature.set(signature, [i]);
        }
    });
    for (const group of bySignature.values()) {
        if (group.length < 2) {
            continue;
        }
        for (const i of group) {
            const related = group.filter((j) => j !== i).map((j) => exceptions[j].rowNumber);
            exceptions[i].reason = 'DUPLICATE';
            exceptions[i].message = duplicateMessage(outstanding[i], related);
            exceptions[i].relatedRowNumbers = related;
        }
    }

    // AMOUNT_MISMATCH: flag only the unambiguous case — exactly one whole unmatched
    // debit vs exactly one whole unmatched credit left in a branch scope, within the
    // band. (FIFO guarantees such a pair sits in different match-key groups.)
    const byGroup = new Map<string, number[]>();
    exceptions.forEach((exc, i) => {
        if (exc.reason === 'DUPLICATE' || outstanding[i].reason.startsWith('PARTIALLY')) {
            return;
        }
        const key = JSON.stringify([exc.entity, exc.gl, exc.branchNumber]);
        const group = byGroup.get(key);
        if (group) {
            group.push(i);
        } else {
            byGroup.set(key, [i]);
        }
    });
    for (const group of byGroup.values()) {
        const debits = group.filter((i) => exceptions[i].direction === 'debit');
        const credits = group.filter((i) => exceptions[i].direction === 'credit');
        if (debits.length !== 1 || credits.length !== 1) {
            continue;
        }
        const d = exceptions[debits[0]];
        const c = exceptions[credits[0]];
        const gap = Math.abs(d.outstandingFils - c.outstandingFils);
        if (gap === 0 || gap > band) {
            continue; // equal amounts would have cleared in F4; beyond the band is not a near-miss
        }
        d.reason = 'AMOUNT_MISMATCH';
        d.message = mismatchMessage(outstanding[debits[0]], outstanding[credits[0]]);
        d.relatedRowNumbers = [c.rowNumber];
        c.reason = 'AMOUNT_MISMATCH';
        c.message = mismatchMessage(outstanding[credits[0]], outstanding[debits[0]]);
        c.relatedRowNumbers = [d.rowNumber];
    }

    const byReason: Partial<Record<LgExceptionReason, number>> = {};
    for (const exc of exceptions) {
        byReason[exc.reason] = (byReason[exc.reason] ?? 0) + 1;
    }

    return { exceptions, summary: { total: exceptions.length, byReason } };
}
