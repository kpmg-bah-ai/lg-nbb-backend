/**
 * GOAL-3 R7 — register-mode exception classification (GOAL-3 §4.6). Pure.
 *
 * Every outstanding item becomes exactly ONE exception — nothing dropped,
 * amounts pass through untouched (GOAL.md §5). Reasons are register-aware:
 *
 *   NON_ISSUANCE_CREDIT           credit with no register issuance (take-on,
 *                                 transfers, redeems) — not a cheque liability
 *   UNRESOLVED_BATCH_DEBIT        batch debit whose Ref.# allocation failed
 *                                 or left a residual
 *   UNMATCHED_LEDGER_DEBIT        debit with no register payment match
 *   REGISTER_LAG_OPS_PAID         ops says PAID, register still unmatched,
 *                                 no ledger evidence — statement excludes it
 *   REGISTER_PAID_NO_LEDGER_DEBIT register says paid; the window lacks the debit
 *   UNMATCHED_CREDIT (base)       an outstanding cheque — a statement line,
 *                                 kept so exceptions ⊇ outstanding
 *
 * On top come informational entries (outstandingFils 0 — they never distort
 * the reconciliation sums): KEY_COLLISION per collision-flagged cheque, and
 * one run-level EXTRACT_GAP when derived ≠ stated.
 */

import {
    ChequeOutcome,
    ExceptionSummary,
    filsToBhd,
    LgException,
    LgExceptionReason,
    OutstandingItem,
} from '../shared/models';
import { RegisterMatchResult } from './registerMatch';

export interface RegisterExceptionResult {
    exceptions: LgException[];
    summary: ExceptionSummary;
}

const fmtBhd = (fils: number): string => filsToBhd(Math.abs(fils)).toFixed(3);

function baseException(item: OutstandingItem, reason: LgExceptionReason, message: string): LgException {
    return {
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
        reason,
        message,
    };
}

function classifyCredit(item: OutstandingItem, outcome: ChequeOutcome | undefined): LgException {
    if (!item.cheque || !outcome) {
        return baseException(
            item,
            'NON_ISSUANCE_CREDIT',
            `This ${fmtBhd(item.outstandingFils)} BHD credit (journal ${item.journalNumber}) has no issuance ` +
                `in the cheque register — an opening take-on, transfer or redeem, not a cheque liability. ` +
                `Reclassify it out of the MC-payable population.`
        );
    }
    const chq = item.cheque.chequeNumber ?? '(unknown)';
    switch (outcome.state) {
        case 'OPS_PAID':
            return baseException(
                item,
                'REGISTER_LAG_OPS_PAID',
                `Cheque #${chq} is marked PAID by operations (journal ${outcome.opsJournal ?? '—'}` +
                    `${outcome.opsDate ? `, ${outcome.opsDate}` : ''}) but the register still shows it ` +
                    `unmatched and no ledger debit was found — a register lag. Feed a status correction ` +
                    `to the cheque system; this is not a reconciliation break.`
            );
        case 'REGISTER_MATCHED_NO_DEBIT':
            return baseException(
                item,
                'REGISTER_PAID_NO_LEDGER_DEBIT',
                `The register shows cheque #${chq} paid on ${outcome.matchedPostDate ?? '—'} ` +
                    `(journal ${outcome.matchedJournal ?? '—'}) but no ledger debit in the window matches — ` +
                    `paid outside the extract window, or the extract is incomplete.`
            );
        case 'STOPPED':
            return baseException(
                item,
                'UNMATCHED_CREDIT',
                `Cheque #${chq} is stopped/cancelled in the register (status ${outcome.status ?? '—'}) with its ` +
                    `${fmtBhd(item.outstandingFils)} BHD issuance credit still uncleared — confirm the reversal entry.`
            );
        default:
            return baseException(
                item,
                'UNMATCHED_CREDIT',
                `Outstanding cheque #${chq}${item.cheque.payee ? ` (payee ${item.cheque.payee})` : ''} — issued ` +
                    `${item.cheque.issuedDate ?? '—'}, ${fmtBhd(item.outstandingFils)} BHD, no payment as at the ` +
                    `review date. A statement line (Section ${item.ageBucket === 'old' ? 'A — Old Items' : 'B'}).`
            );
    }
}

function classifyDebit(item: OutstandingItem): LgException {
    if (item.batchRefs && item.batchRefs.length > 0) {
        const residual =
            item.outstandingFils < item.originalFils
                ? `${fmtBhd(item.outstandingFils)} BHD of ${fmtBhd(item.originalFils)} BHD remains unallocated`
                : `none of its ${fmtBhd(item.originalFils)} BHD could be allocated`;
        return baseException(
            item,
            'UNRESOLVED_BATCH_DEBIT',
            `Batch debit (journal ${item.journalNumber}) references Ref.# ${item.batchRefs.join(', ')} but ` +
                `${residual} against register cheques — verify the referenced journals and amounts.`
        );
    }
    const note = item.reconciledNote ? ` Manually dispositioned: "${item.reconciledNote}".` : '';
    return baseException(
        item,
        'UNMATCHED_LEDGER_DEBIT',
        `This ${fmtBhd(item.outstandingFils)} BHD debit (journal ${item.journalNumber}) matches no register ` +
            `payment key and carries no Ref.# list — a miscoded or non-cheque entry to investigate.${note}`
    );
}

/** Classifies a register-mode match result into reviewer-facing exceptions. */
export function classifyRegisterExceptions(
    match: RegisterMatchResult,
    extractGapFils: number | undefined
): RegisterExceptionResult {
    const outcomeByRow = new Map(match.outcomes.map((o) => [o.rowNumber, o]));

    const exceptions: LgException[] = match.outstanding.map((item) =>
        item.direction === 'credit'
            ? classifyCredit(item, item.cheque ? outcomeByRow.get(item.cheque.registerRowNumber) : undefined)
            : classifyDebit(item)
    );

    // Informational: legitimate key collisions — handled one-for-one, shown anyway.
    const entity = match.outstanding[0]?.entity ?? '';
    const gl = match.outstanding[0]?.gl ?? '';
    for (const outcome of match.outcomes) {
        if (!outcome.keyCollision) {
            continue;
        }
        exceptions.push({
            entity,
            gl,
            branchNumber: outcome.issuedBranch ?? '',
            postDate: outcome.issuedPostDate ?? outcome.issuedDate ?? '',
            direction: 'credit',
            originalFils: outcome.amountFils,
            outstandingFils: 0, // informational — never distorts the sums
            journalNumber: outcome.issuedJournal ?? '',
            rowNumber: outcome.rowNumber,
            sheet: outcome.sheet,
            ageBucket: outcome.ageBucket ?? 'current',
            reason: 'KEY_COLLISION',
            message:
                `Cheque #${outcome.chequeNumber ?? '(unknown)'} shares its (date, journal, amount) key with ` +
                `other instruments or ledger rows — multiset matching paired occurrences one-for-one; ` +
                `verify the allocation (state: ${outcome.state}).`,
        });
    }

    // Run-level: the ledger extract does not tie to the stated EoD balance.
    if (extractGapFils !== undefined && extractGapFils !== 0) {
        exceptions.push({
            entity,
            gl,
            branchNumber: '',
            postDate: match.summary.asOf,
            direction: extractGapFils < 0 ? 'credit' : 'debit',
            originalFils: Math.abs(extractGapFils),
            outstandingFils: 0, // informational — the gap is not a posting
            journalNumber: '',
            rowNumber: 0,
            ageBucket: 'current',
            reason: 'EXTRACT_GAP',
            message:
                `The balance derived from the ledger rows differs from the stated End Date EoD Balance by ` +
                `${fmtBhd(extractGapFils)} BHD — the extract is missing movements (or the stated balance ` +
                `covers items outside these sheets). Request a complete extract to close the gap.`,
        });
    }

    const byReason: Partial<Record<LgExceptionReason, number>> = {};
    for (const exc of exceptions) {
        byReason[exc.reason] = (byReason[exc.reason] ?? 0) + 1;
    }
    return { exceptions, summary: { total: exceptions.length, byReason } };
}
