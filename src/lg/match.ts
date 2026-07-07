/**
 * LG reconciliation — debit↔credit matching engine (GOAL.md §3, §4 F4). Pure functions.
 *
 * Model (GOAL.md §11.3): postings are grouped by a match key; within a group, debits
 * and credits offset each other by amount, FIFO in posting order. A group whose signed
 * fils sum to zero is a fully cleared set — this uniformly covers 1:1, one-to-many
 * (an 11k credit cleared by 2k + 9k debits) and many-to-many. Whatever cannot be
 * offset remains as OutstandingItem fragments: whole postings (UNMATCHED_*) or the
 * residual of partially cleared ones (PARTIALLY_MATCHED_*, e.g. 2k left of an 11k
 * credit offset by only 9k). Nothing is ever netted silently — every fragment keeps
 * its source row for drill-down (GOAL.md §5 traceability).
 *
 * Invariant (tested): Σ signed outstanding == Σ signed postings == the F3 GL balance,
 * which is exactly the F5 identity `GL Balance − Σ(outstanding) = 0` at the engine level.
 *
 * The match key defaults to (entity, gl, branchNumber, accountNumber) — GOAL.md §9.2
 * is still open, so the key is a parameter. Aging: a fragment older than one year at
 * the review date is an "Old Item", otherwise current ("Less than 1 year", §9.4).
 */

import {
    AgeBucket,
    filsToBhd,
    MatchedLeg,
    MatchedSet,
    MatchSummary,
    OutstandingItem,
    ParsedPosting,
} from '../shared/models';
import { deriveAsOf } from './balance';

export const DEFAULT_MATCH_KEY: (keyof ParsedPosting)[] = ['entity', 'gl', 'branchNumber', 'accountNumber'];

/** Fragments strictly older than this many days at asOf are "Old Items". */
export const OLD_AFTER_DAYS = 365;

export interface MatchOptions {
    /** Review date (yyyy-mm-dd); defaults to the latest post date in the data. */
    asOf?: string;
    matchKey?: (keyof ParsedPosting)[];
    oldAfterDays?: number;
}

export interface MatchResult {
    outstanding: OutstandingItem[];
    /** Cleared sets — how each instrument/chain actually cleared (GOAL-2 G1). */
    matchedSets: MatchedSet[];
    summary: MatchSummary;
}

const MS_PER_DAY = 86_400_000;

function daysBetween(fromIso: string, toIso: string): number {
    return Math.floor((Date.parse(toIso) - Date.parse(fromIso)) / MS_PER_DAY);
}

interface Leg {
    posting: ParsedPosting;
    remainingFils: number;
}

function toOutstanding(leg: Leg, asOf: string, oldAfterDays: number): OutstandingItem {
    const p = leg.posting;
    const originalFils = Math.abs(p.amountBhdFils);
    const partial = leg.remainingFils < originalFils;
    const ageBucket: AgeBucket = daysBetween(p.postDate, asOf) > oldAfterDays ? 'old' : 'current';
    const reason =
        p.direction === 'debit'
            ? partial
                ? 'PARTIALLY_MATCHED_DEBIT'
                : 'UNMATCHED_DEBIT'
            : partial
            ? 'PARTIALLY_MATCHED_CREDIT'
            : 'UNMATCHED_CREDIT';
    return {
        entity: p.entity,
        gl: p.gl,
        branchNumber: p.branchNumber,
        accountNumber: p.accountNumber,
        postDate: p.postDate,
        direction: p.direction,
        originalFils,
        outstandingFils: leg.remainingFils,
        outstanding: filsToBhd(leg.remainingFils),
        logCode: p.logCode,
        journalNumber: p.journalNumber,
        sequence: p.sequence,
        rowNumber: p.rowNumber,
        sheet: p.sheet,
        ageBucket,
        reason,
    };
}

function toMatchedLeg(leg: Leg, matchedFils: number): MatchedLeg {
    const p = leg.posting;
    return {
        postDate: p.postDate,
        direction: p.direction,
        originalFils: Math.abs(p.amountBhdFils),
        matchedFils,
        journalNumber: p.journalNumber,
        sequence: p.sequence,
        logCode: p.logCode,
        rowNumber: p.rowNumber,
        sheet: p.sheet,
    };
}

/** Pairs debits with credits per match-key group; returns the cleared sets and what could not be cleared. */
export function matchPostings(postings: ParsedPosting[], options: MatchOptions = {}): MatchResult {
    const matchKey = options.matchKey ?? DEFAULT_MATCH_KEY;
    const oldAfterDays = options.oldAfterDays ?? OLD_AFTER_DAYS;
    const asOf = options.asOf ?? deriveAsOf(postings) ?? '1970-01-01';

    // F4 must operate on the same population as F3: postings after the review date do
    // not exist yet for this reconciliation. Without this, a back-dated review would
    // let a future credit silently clear a past debit and break the tie-out identity.
    const included = postings.filter((p) => p.postDate <= asOf);

    const groups = new Map<string, ParsedPosting[]>();
    for (const p of included) {
        const key = JSON.stringify(matchKey.map((field) => p[field] ?? ''));
        const group = groups.get(key);
        if (group) {
            group.push(p);
        } else {
            groups.set(key, [p]);
        }
    }

    const outstanding: OutstandingItem[] = [];
    const matchedSets: MatchedSet[] = [];
    let matchedFils = 0;

    for (const group of groups.values()) {
        const sorted = [...group].sort((a, b) =>
            a.postDate < b.postDate ? -1 : a.postDate > b.postDate ? 1 : a.rowNumber - b.rowNumber
        );
        const debits: Leg[] = [];
        const credits: Leg[] = [];
        for (const p of sorted) {
            const fils = Math.abs(p.amountBhdFils);
            if (fils === 0) {
                continue; // zero-amount legs (direction from log code) clear nothing and owe nothing
            }
            (p.direction === 'debit' ? debits : credits).push({ posting: p, remainingFils: fils });
        }

        // FIFO offset: consume the oldest debit against the oldest credit, splitting
        // whichever is larger — this is what makes 11k = 2k + 9k a full clear. Each
        // offset is recorded so cleared sets can be reconstructed below (G1).
        const offsets: { d: number; c: number; fils: number }[] = [];
        let d = 0;
        let c = 0;
        while (d < debits.length && c < credits.length) {
            const offset = Math.min(debits[d].remainingFils, credits[c].remainingFils);
            debits[d].remainingFils -= offset;
            credits[c].remainingFils -= offset;
            matchedFils += offset;
            offsets.push({ d, c, fils: offset });
            if (debits[d].remainingFils === 0) {
                d++;
            }
            if (credits[c].remainingFils === 0) {
                c++;
            }
        }

        // Cleared sets: connected components of the offset graph. FIFO advances d and c
        // monotonically, so a new offset either shares a leg with the running component
        // (same debit or same credit still being consumed) or — when the previous offset
        // exhausted both legs at once — starts a fresh component.
        const components: { offsets: typeof offsets; debitIdx: Set<number>; creditIdx: Set<number> }[] = [];
        for (const off of offsets) {
            const current = components[components.length - 1];
            if (current && (current.debitIdx.has(off.d) || current.creditIdx.has(off.c))) {
                current.offsets.push(off);
                current.debitIdx.add(off.d);
                current.creditIdx.add(off.c);
            } else {
                components.push({ offsets: [off], debitIdx: new Set([off.d]), creditIdx: new Set([off.c]) });
            }
        }

        for (const component of components) {
            const legMatched = new Map<string, number>();
            let setFils = 0;
            for (const off of component.offsets) {
                setFils += off.fils;
                legMatched.set(`d${off.d}`, (legMatched.get(`d${off.d}`) ?? 0) + off.fils);
                legMatched.set(`c${off.c}`, (legMatched.get(`c${off.c}`) ?? 0) + off.fils);
            }
            const debitLegs = [...component.debitIdx]
                .sort((a, b) => a - b)
                .map((i) => toMatchedLeg(debits[i], legMatched.get(`d${i}`) ?? 0));
            const creditLegs = [...component.creditIdx]
                .sort((a, b) => a - b)
                .map((i) => toMatchedLeg(credits[i], legMatched.get(`c${i}`) ?? 0));
            const fullyCleared =
                [...component.debitIdx].every((i) => debits[i].remainingFils === 0) &&
                [...component.creditIdx].every((i) => credits[i].remainingFils === 0);
            const participants = [
                ...[...component.debitIdx].map((i) => debits[i].posting),
                ...[...component.creditIdx].map((i) => credits[i].posting),
            ];
            const accounts = new Set(participants.map((p) => p.accountNumber));
            const firstCreditDate = creditLegs.reduce(
                (min, l) => (l.postDate < min ? l.postDate : min),
                creditLegs[0].postDate
            );
            const finalDebitDate = debitLegs.reduce(
                (max, l) => (l.postDate > max ? l.postDate : max),
                debitLegs[0].postDate
            );
            matchedSets.push({
                entity: participants[0].entity,
                gl: participants[0].gl,
                branchNumber: participants[0].branchNumber,
                accountNumber: accounts.size === 1 ? participants[0].accountNumber : undefined,
                matchedFils: setFils,
                creditLegs,
                debitLegs,
                firstCreditDate,
                finalDebitDate,
                settledDays: daysBetween(firstCreditDate, finalDebitDate),
                fullyCleared,
            });
        }

        for (const leg of [...debits, ...credits]) {
            if (leg.remainingFils > 0) {
                outstanding.push(toOutstanding(leg, asOf, oldAfterDays));
            }
        }
    }

    outstanding.sort(
        (a, b) =>
            a.branchNumber.localeCompare(b.branchNumber) ||
            (a.postDate < b.postDate ? -1 : a.postDate > b.postDate ? 1 : a.rowNumber - b.rowNumber)
    );
    matchedSets.sort(
        (a, b) =>
            a.branchNumber.localeCompare(b.branchNumber) ||
            (a.firstCreditDate < b.firstCreditDate
                ? -1
                : a.firstCreditDate > b.firstCreditDate
                ? 1
                : a.creditLegs[0].rowNumber - b.creditLegs[0].rowNumber)
    );

    let outstandingDebitFils = 0;
    let outstandingCreditFils = 0;
    let oldCount = 0;
    const byBranch = new Map<
        string,
        { branchNumber: string; outstandingCount: number; outstandingFils: number; matchedSetCount: number }
    >();
    const branchEntry = (branchNumber: string) => {
        let branch = byBranch.get(branchNumber);
        if (!branch) {
            branch = { branchNumber, outstandingCount: 0, outstandingFils: 0, matchedSetCount: 0 };
            byBranch.set(branchNumber, branch);
        }
        return branch;
    };
    for (const item of outstanding) {
        if (item.direction === 'debit') {
            outstandingDebitFils += item.outstandingFils;
        } else {
            outstandingCreditFils += item.outstandingFils;
        }
        if (item.ageBucket === 'old') {
            oldCount++;
        }
        const branch = branchEntry(item.branchNumber);
        branch.outstandingCount++;
        branch.outstandingFils += item.direction === 'debit' ? item.outstandingFils : -item.outstandingFils;
    }
    for (const set of matchedSets) {
        branchEntry(set.branchNumber).matchedSetCount++;
    }

    const summary: MatchSummary = {
        asOf,
        matchKey: matchKey.map(String),
        matchedFils,
        outstandingCount: outstanding.length,
        outstandingDebitFils,
        outstandingCreditFils,
        netOutstandingFils: outstandingDebitFils - outstandingCreditFils,
        oldCount,
        currentCount: outstanding.length - oldCount,
        matchedSetCount: matchedSets.length,
        fullyClearedSetCount: matchedSets.filter((s) => s.fullyCleared).length,
        byBranch: [...byBranch.values()].sort((a, b) => a.branchNumber.localeCompare(b.branchNumber)),
    };

    return { outstanding, matchedSets, summary };
}
