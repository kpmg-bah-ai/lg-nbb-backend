/**
 * GOAL-3 R4/R5 — two-legged GL↔register matching (GOAL-3 §4.4). Pure functions.
 *
 * The register is the source of truth for instrument identity: a cheque's
 * issuance is the GL CREDIT whose (transaction date, journal, |amount|) equals
 * the register's issued (post date, journal, amount); its payment is the GL
 * DEBIT matching the register's matched (post date, journal, amount). Keys are
 * TUPLES compared structurally — never string concatenation — and buckets are
 * MULTISETS consumed one-for-one FIFO by row number, because the reference file
 * proves the key legitimately collides (34 register keys / 108 debit keys).
 *
 * Matching is identification; only credit↔debit pairs net. A MatchedSet forms
 * ONLY when both legs are in the ledger window, so every set nets to zero and
 * the engine invariant survives verbatim (tested):
 *     Σ signed outstanding = Σ signed included postings = derived balance.
 *
 * Passes (each consumes matched occurrences):
 *   1. issuance — GL credits vs register issued keys (transaction date);
 *   2. payment  — GL debits  vs register matched keys (transaction date);
 *   3. variant  — leftovers retried with the POSTING date (the reference file
 *      carries 2 credits whose transaction date drifts; flagged);
 *   4. batch    — unmatched DEBIT POSTING rows resolved via their Detailed
 *      Description `Ref.#` journal lists (Task 7);
 *   5. ops      — reviewer PAID dispositions exclude cheques from the
 *      outstanding STATEMENT while their GL credit stays outstanding (Task 7).
 *
 * Aging: outstanding cheques age from the register issuance date (§9.4);
 * non-cheque postings keep post-date aging.
 */

import {
    AgeBucket,
    ChequeAttributes,
    ChequeOutcome,
    filsToBhd,
    MatchedLeg,
    MatchedSet,
    MatchedVia,
    MatchSummary,
    OutstandingItem,
    ParsedPosting,
    RegisterCheque,
} from '../shared/models';
import { deriveAsOf } from './balance';
import { OLD_AFTER_DAYS } from './match';

/** The §9.2 answer: the tuple key per leg. */
export const REGISTER_MATCH_KEY = ['transactionDate', 'journalNumber', 'amountFils'];

/**
 * Extracts the `Ref.#` journal list a batch DEBIT POSTING embeds in its
 * Detailed Description (file §5.3) — e.g. "DEBIT POSTING-20-Ref.# 100242402,…".
 * Deduped, order-preserving. Tolerates spacing/punctuation drift.
 */
export function parseBatchRefs(text: string | undefined): string[] {
    if (!text) {
        return [];
    }
    const refs: string[] = [];
    const seen = new Set<string>();
    const pattern = /Ref\s*[.．]?\s*#\s*(\d+)/gi;
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(text)) !== null) {
        if (!seen.has(match[1])) {
            seen.add(match[1]);
            refs.push(match[1]);
        }
    }
    return refs;
}

const MS_PER_DAY = 86_400_000;

function daysBetween(fromIso: string, toIso: string): number {
    return Math.floor((Date.parse(toIso) - Date.parse(fromIso)) / MS_PER_DAY);
}

/** Structural tuple key — JSON of (dateIso, journal, |fils|), never concatenation. */
function keyOf(dateIso: string, journalNumber: string, amountFils: number): string {
    return JSON.stringify([dateIso, journalNumber, amountFils]);
}

export interface RegisterMatchOptions {
    /** Review date (yyyy-mm-dd); defaults to the latest post date in the data. */
    asOf?: string;
    oldAfterDays?: number;
}

export interface RegisterMatchResult {
    outstanding: OutstandingItem[];
    matchedSets: MatchedSet[];
    /** One outcome per register cheque — the instrument-level story. */
    outcomes: ChequeOutcome[];
    summary: MatchSummary;
}

interface LegLink {
    postingIdx: number;
    via: MatchedVia;
}

function toMatchedLeg(p: ParsedPosting): MatchedLeg {
    return {
        postDate: p.postDate,
        direction: p.direction,
        originalFils: Math.abs(p.amountBhdFils),
        matchedFils: Math.abs(p.amountBhdFils),
        journalNumber: p.journalNumber,
        sequence: p.sequence,
        logCode: p.logCode,
        rowNumber: p.rowNumber,
        sheet: p.sheet,
    };
}

function toChequeAttributes(c: RegisterCheque): ChequeAttributes {
    return {
        chequeNumber: c.chequeNumber,
        payee: c.payee,
        issuedDate: c.issuedDate,
        status: c.status,
        purchaser: c.purchaser,
        opsRemark: c.opsRemark,
        registerRowNumber: c.rowNumber,
    };
}

/** Multiset buckets: key → cheque indexes, FIFO by register row number. */
function buildBuckets(
    cheques: RegisterCheque[],
    dateOf: (c: RegisterCheque) => string | undefined,
    journalOf: (c: RegisterCheque) => string | undefined
): Map<string, number[]> {
    const buckets = new Map<string, number[]>();
    cheques
        .map((c, i) => i)
        .sort((a, b) => cheques[a].rowNumber - cheques[b].rowNumber)
        .forEach((i) => {
            const date = dateOf(cheques[i]);
            const journal = journalOf(cheques[i]);
            if (!date || !journal) {
                return;
            }
            const key = keyOf(date, journal, cheques[i].amountFils);
            const bucket = buckets.get(key);
            if (bucket) {
                bucket.push(i);
            } else {
                buckets.set(key, [i]);
            }
        });
    return buckets;
}

/** Pairs debits with credits ACROSS the register per GOAL-3 §4.4; see module doc. */
export function matchRegister(
    postings: ParsedPosting[],
    cheques: RegisterCheque[],
    options: RegisterMatchOptions = {}
): RegisterMatchResult {
    const oldAfterDays = options.oldAfterDays ?? OLD_AFTER_DAYS;
    const asOf = options.asOf ?? deriveAsOf(postings) ?? '1970-01-01';

    // Same population rule as the FIFO engine: postings after the review date
    // do not exist for this reconciliation (GOAL.md §11.3).
    const included = postings
        .map((p, i) => i)
        .filter((i) => postings[i].postDate <= asOf)
        .sort((a, b) => postings[a].rowNumber - postings[b].rowNumber);
    const credits = included.filter((i) => postings[i].direction === 'credit');
    const debits = included.filter((i) => postings[i].direction === 'debit');

    const issuanceBuckets = buildBuckets(
        cheques,
        (c) => c.issuedPostDate,
        (c) => c.issuedJournal
    );
    const paymentBuckets = buildBuckets(
        cheques,
        (c) => c.matchedPostDate,
        (c) => c.matchedJournal
    );

    // Ledger-side key multiplicity (108 colliding debit keys in the reference file).
    const ledgerKeyCount = new Map<string, number>();
    for (const i of included) {
        const p = postings[i];
        const date = p.transactionDate ?? p.postDate;
        const key = keyOf(date, p.journalNumber, Math.abs(p.amountBhdFils));
        ledgerKeyCount.set(key, (ledgerKeyCount.get(key) ?? 0) + 1);
    }

    const issuanceOf: (LegLink | undefined)[] = new Array(cheques.length).fill(undefined);
    const paymentOf: (LegLink | undefined)[] = new Array(cheques.length).fill(undefined);
    const postingConsumed: boolean[] = new Array(postings.length).fill(false);

    const runPass = (
        postingIdxs: number[],
        dateOf: (p: ParsedPosting) => string | undefined,
        buckets: Map<string, number[]>,
        linkOf: (LegLink | undefined)[],
        via: MatchedVia
    ): void => {
        for (const idx of postingIdxs) {
            if (postingConsumed[idx]) {
                continue;
            }
            const p = postings[idx];
            const date = dateOf(p);
            if (!date) {
                continue;
            }
            const bucket = buckets.get(keyOf(date, p.journalNumber, Math.abs(p.amountBhdFils)));
            if (!bucket) {
                continue;
            }
            // Multiset one-for-one: the first cheque occurrence whose leg is still free.
            const chequeIdx = bucket.find((c) => linkOf[c] === undefined);
            if (chequeIdx === undefined) {
                continue;
            }
            linkOf[chequeIdx] = { postingIdx: idx, via };
            postingConsumed[idx] = true;
        }
    };

    // Pass 1 + 2: the tuple key on the transaction date.
    runPass(credits, (p) => p.transactionDate ?? p.postDate, issuanceBuckets, issuanceOf, 'KEY');
    runPass(debits, (p) => p.transactionDate ?? p.postDate, paymentBuckets, paymentOf, 'KEY');
    // Pass 3: posting-date variant for legs whose transaction date drifted.
    const variantDate = (p: ParsedPosting) =>
        p.transactionDate !== undefined && p.transactionDate !== p.postDate ? p.postDate : undefined;
    runPass(credits, variantDate, issuanceBuckets, issuanceOf, 'POSTING_DATE_VARIANT');
    runPass(debits, variantDate, paymentBuckets, paymentOf, 'POSTING_DATE_VARIANT');

    // Pass 4 (GOAL-3 §4.4): batch debits — one debit paying many cheques, its
    // Detailed Description embedding the Ref.# journals of what it pays. FIFO
    // by issue date, allocating only while Σ cheque amounts ≤ the debit amount:
    // a batch can never clear more than its own value.
    interface BatchAllocation {
        debitIdx: number;
        chequeIdxs: number[];
        allocatedFils: number;
    }
    const batchAllocations: BatchAllocation[] = [];
    const batchRefsByDebit = new Map<number, string[]>();
    for (const idx of debits) {
        if (postingConsumed[idx]) {
            continue;
        }
        const refs = parseBatchRefs(postings[idx].detailedDescription);
        if (refs.length === 0) {
            continue;
        }
        batchRefsByDebit.set(idx, refs);
        const refSet = new Set(refs);
        const debitFils = Math.abs(postings[idx].amountBhdFils);
        const candidates = cheques
            .map((_, i) => i)
            .filter(
                (i) =>
                    issuanceOf[i] !== undefined &&
                    paymentOf[i] === undefined &&
                    ((cheques[i].issuedJournal !== undefined && refSet.has(cheques[i].issuedJournal!)) ||
                        (cheques[i].opsJournal !== undefined && refSet.has(cheques[i].opsJournal!)))
            )
            .sort((a, b) => {
                const dateA = cheques[a].issuedDate ?? cheques[a].issuedPostDate ?? '';
                const dateB = cheques[b].issuedDate ?? cheques[b].issuedPostDate ?? '';
                return dateA < dateB ? -1 : dateA > dateB ? 1 : cheques[a].rowNumber - cheques[b].rowNumber;
            });
        const chequeIdxs: number[] = [];
        let allocatedFils = 0;
        for (const i of candidates) {
            if (allocatedFils + cheques[i].amountFils <= debitFils) {
                chequeIdxs.push(i);
                allocatedFils += cheques[i].amountFils;
            }
        }
        if (chequeIdxs.length === 0) {
            continue;
        }
        for (const i of chequeIdxs) {
            paymentOf[i] = { postingIdx: idx, via: 'BATCH_REF' };
        }
        postingConsumed[idx] = true;
        batchAllocations.push({ debitIdx: idx, chequeIdxs, allocatedFils });
    }

    // Collision visibility: a cheque whose key bucket (either leg) or ledger key
    // had more than one occupant — handled correctly, but reviewers must see it.
    const collides = (c: RegisterCheque): boolean => {
        const issKey = c.issuedPostDate && c.issuedJournal ? keyOf(c.issuedPostDate, c.issuedJournal, c.amountFils) : undefined;
        const payKey = c.matchedPostDate && c.matchedJournal ? keyOf(c.matchedPostDate, c.matchedJournal, c.amountFils) : undefined;
        return (
            (issKey !== undefined &&
                ((issuanceBuckets.get(issKey)?.length ?? 0) > 1 || (ledgerKeyCount.get(issKey) ?? 0) > 1)) ||
            (payKey !== undefined &&
                ((paymentBuckets.get(payKey)?.length ?? 0) > 1 || (ledgerKeyCount.get(payKey) ?? 0) > 1))
        );
    };

    // Matched sets: only cheques with BOTH legs in the window net to zero.
    const matchedSets: MatchedSet[] = [];
    const inSet: boolean[] = new Array(postings.length).fill(false);
    /** Batch debits only partly consumed: posting idx → allocated fils. */
    const partialAllocated = new Map<number, number>();
    let matchedFils = 0;
    cheques.forEach((c, i) => {
        const issuance = issuanceOf[i];
        const payment = paymentOf[i];
        if (!issuance || !payment || payment.via === 'BATCH_REF') {
            return; // batch payments form one set per DEBIT below
        }
        const credit = postings[issuance.postingIdx];
        const debit = postings[payment.postingIdx];
        inSet[issuance.postingIdx] = true;
        inSet[payment.postingIdx] = true;
        matchedFils += c.amountFils;
        const sameAccount = credit.accountNumber === debit.accountNumber;
        matchedSets.push({
            entity: credit.entity,
            gl: credit.gl,
            branchNumber: credit.branchNumber,
            accountNumber: sameAccount ? credit.accountNumber : undefined,
            matchedFils: c.amountFils,
            creditLegCount: 1,
            debitLegCount: 1,
            creditLegs: [toMatchedLeg(credit)],
            debitLegs: [toMatchedLeg(debit)],
            firstCreditDate: credit.postDate,
            finalDebitDate: debit.postDate,
            settledDays: daysBetween(credit.postDate, debit.postDate),
            fullyCleared: true,
            chequeNumber: c.chequeNumber,
            matchedVia: issuance.via === 'POSTING_DATE_VARIANT' || payment.via === 'POSTING_DATE_VARIANT'
                ? 'POSTING_DATE_VARIANT'
                : payment.via,
        });
    });

    // Batch sets: one per batch debit — N issuance credit legs against one debit
    // leg consumed up to the allocated amount. A partial allocation leaves the
    // residual outstanding (never netted silently); the set is not fullyCleared.
    for (const batch of batchAllocations) {
        const debit = postings[batch.debitIdx];
        const debitFils = Math.abs(debit.amountBhdFils);
        const creditPostings = batch.chequeIdxs.map((i) => postings[issuanceOf[i]!.postingIdx]);
        for (const i of batch.chequeIdxs) {
            inSet[issuanceOf[i]!.postingIdx] = true;
        }
        const fullyCleared = batch.allocatedFils === debitFils;
        if (fullyCleared) {
            inSet[batch.debitIdx] = true;
        } else {
            partialAllocated.set(batch.debitIdx, batch.allocatedFils);
        }
        matchedFils += batch.allocatedFils;
        const firstCreditDate = creditPostings.reduce(
            (min, p) => (p.postDate < min ? p.postDate : min),
            creditPostings[0].postDate
        );
        matchedSets.push({
            entity: debit.entity,
            gl: debit.gl,
            branchNumber: debit.branchNumber,
            accountNumber: undefined,
            matchedFils: batch.allocatedFils,
            creditLegCount: creditPostings.length,
            debitLegCount: 1,
            creditLegs: creditPostings.map(toMatchedLeg),
            debitLegs: [{ ...toMatchedLeg(debit), matchedFils: batch.allocatedFils }],
            firstCreditDate,
            finalDebitDate: debit.postDate,
            settledDays: daysBetween(firstCreditDate, debit.postDate),
            fullyCleared,
            chequeNumber: batch.chequeIdxs.length === 1 ? cheques[batch.chequeIdxs[0]].chequeNumber : undefined,
            matchedVia: 'BATCH_REF',
        });
    }

    // Cheque attributes for issuance-linked credits that did NOT clear.
    const chequeOfCredit = new Map<number, number>(); // posting idx → cheque idx
    issuanceOf.forEach((link, chequeIdx) => {
        if (link && !inSet[link.postingIdx]) {
            chequeOfCredit.set(link.postingIdx, chequeIdx);
        }
    });

    // Outstanding: every included posting not inside a zero-net set. Partially
    // allocated batch debits surface their residual — never netted silently.
    const outstanding: OutstandingItem[] = [];
    for (const idx of included) {
        if (inSet[idx]) {
            continue;
        }
        const p = postings[idx];
        const fils = Math.abs(p.amountBhdFils);
        if (fils === 0) {
            continue;
        }
        const allocated = partialAllocated.get(idx);
        const outstandingFils = allocated !== undefined ? fils - allocated : fils;
        const chequeIdx = chequeOfCredit.get(idx);
        const cheque = chequeIdx !== undefined ? cheques[chequeIdx] : undefined;
        // Cheque-backed credits age from the register issuance date (§9.4).
        const ageDate = cheque?.issuedDate ?? cheque?.issuedPostDate ?? p.postDate;
        const ageBucket: AgeBucket = daysBetween(ageDate, asOf) > oldAfterDays ? 'old' : 'current';
        outstanding.push({
            entity: p.entity,
            gl: p.gl,
            branchNumber: p.branchNumber,
            accountNumber: p.accountNumber,
            postDate: p.postDate,
            direction: p.direction,
            originalFils: fils,
            outstandingFils,
            outstanding: filsToBhd(outstandingFils),
            logCode: p.logCode,
            journalNumber: p.journalNumber,
            sequence: p.sequence,
            rowNumber: p.rowNumber,
            sheet: p.sheet,
            ageBucket,
            reason:
                p.direction === 'debit'
                    ? allocated !== undefined
                        ? 'PARTIALLY_MATCHED_DEBIT'
                        : 'UNMATCHED_DEBIT'
                    : 'UNMATCHED_CREDIT',
            cheque: cheque ? toChequeAttributes(cheque) : undefined,
            batchRefs: batchRefsByDebit.get(idx),
            reconciledNote: p.reconciledNote,
        });
    }

    // Instrument-level outcomes (GOAL-3 §4.4 precedence).
    const outcomes: ChequeOutcome[] = cheques.map((c, i) => {
        const issuance = issuanceOf[i];
        const payment = paymentOf[i];
        const registerMatched = c.matchedPostDate !== undefined || c.matchedJournal !== undefined;
        let state: ChequeOutcome['state'];
        if (!issuance) {
            state = 'PRE_WINDOW';
        } else if (payment) {
            state = payment.via === 'BATCH_REF' ? 'PAID_VIA_BATCH' : 'PAID';
        } else if (registerMatched) {
            state = 'REGISTER_MATCHED_NO_DEBIT';
        } else if (c.status === '04' && (c.stopReason !== undefined || c.cancelDate !== undefined)) {
            state = 'STOPPED';
        } else if (c.opsPaid) {
            // Pass 5: the reviewer says paid but the ledger carries no evidence —
            // excluded from the outstanding STATEMENT, classified as register lag;
            // the GL credit itself stays outstanding (invariant).
            state = 'OPS_PAID';
        } else {
            state = 'OUTSTANDING';
        }
        const ageDate = c.issuedDate ?? c.issuedPostDate;
        return {
            ...c,
            state,
            matchedVia: payment
                ? payment.via === 'BATCH_REF'
                    ? 'BATCH_REF'
                    : issuance && issuance.via === 'POSTING_DATE_VARIANT'
                    ? 'POSTING_DATE_VARIANT'
                    : payment.via
                : undefined,
            ageBucket:
                state === 'OUTSTANDING' && ageDate
                    ? daysBetween(ageDate, asOf) > oldAfterDays
                        ? 'old'
                        : 'current'
                    : undefined,
            keyCollision: collides(c) || undefined,
            issuanceRowNumber: issuance ? postings[issuance.postingIdx].rowNumber : undefined,
            paymentRowNumbers: payment ? [postings[payment.postingIdx].rowNumber] : undefined,
        };
    });

    // Stamp each outstanding cheque item with its instrument's outcome state so
    // stored items can drive the statement filter without re-deriving it.
    const stateByRegisterRow = new Map(outcomes.map((o) => [o.rowNumber, o.state]));
    for (const item of outstanding) {
        if (item.cheque) {
            item.cheque.state = stateByRegisterRow.get(item.cheque.registerRowNumber);
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
        matchKey: [...REGISTER_MATCH_KEY],
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

    return { outstanding, matchedSets, outcomes, summary };
}

/**
 * The outstanding-items STATEMENT population (GOAL-3 §4.5): cheque-backed
 * outstanding credits whose instrument state is OUTSTANDING. Ops-PAID,
 * register-matched-without-debit and non-issuance items stay in the engine's
 * outstanding set (invariant) but belong to the exceptions decomposition,
 * not the Section A/B statement lines.
 */
export function statementOutstanding(result: RegisterMatchResult): OutstandingItem[] {
    const stateByRow = new Map(result.outcomes.map((o) => [o.rowNumber, o.state]));
    return result.outstanding.filter(
        (o) => o.cheque !== undefined && stateByRow.get(o.cheque.registerRowNumber) === 'OUTSTANDING'
    );
}
