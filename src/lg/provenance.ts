/**
 * GOAL-5 — number provenance.
 *
 * "For each number there should be a description/assessment of how we got it / why,
 * and a nice explanation for it." This module turns a finished run into a list of
 * `ExplainedFigure`s: each headline number the app reports, paired with a `basis`
 * (HOW it was derived — the formula or source) and an `assessment` (WHY it matters
 * and what a reviewer should conclude). Stored on the run and rendered beside the
 * figure, so no number is ever shown bare.
 *
 * It only narrates figures the engine already computed (reconcile / registerReconcile
 * / match / sheetBalances) — it never re-derives money, so the explanations tie to
 * the screen by construction (GOAL.md §5). Pure and deterministic.
 */

import {
    BranchBalance,
    ChequeState,
    ExceptionSummary,
    ExplainedFigure,
    filsToBhd,
    LgRunMode,
    MatchSummary,
    ParseSummary,
    Reconciliation,
    SheetBalance,
} from '../shared/models';
import { fmtBhd } from './export';

export interface ExplainInput {
    mode: LgRunMode;
    summary: ParseSummary;
    asOf?: string;
    balances: BranchBalance[];
    sheetBalances: SheetBalance[];
    reconciliation?: Reconciliation;
    matching?: MatchSummary;
    exceptionsSummary?: ExceptionSummary;
    chequeCount?: number;
    chequesByState?: Partial<Record<ChequeState, number>>;
}

/** Signed fils → display magnitude with sign (reuses the export formatter). */
function money(fils: number): string {
    return `BHD ${fmtBhd(fils)}`;
}

interface FigureSpec {
    key: string;
    label: string;
    group: ExplainedFigure['group'];
    basis: string;
    assessment: string;
    inputs?: string[];
    sheet?: string;
    flag?: boolean;
}

function moneyFigure(fils: number, spec: FigureSpec): ExplainedFigure {
    return { ...spec, valueFils: fils, display: money(fils) };
}

function countFigure(n: number, spec: FigureSpec): ExplainedFigure {
    return { ...spec, count: n, display: n.toLocaleString() };
}

export function explainRun(input: ExplainInput): ExplainedFigure[] {
    const figures: ExplainedFigure[] = [];
    const { summary, sheetBalances } = input;

    // ── Input population ──────────────────────────────────────────────────────
    // In register mode the register sheet's rows parse into cheques, not postings,
    // so they are "handled" even though they are not GL postings — don't count them
    // as parse failures.
    const chequeRows = input.chequeCount ?? 0;
    const handled = summary.parsed + chequeRows;
    const failed = Math.max(0, summary.dataRows - handled);
    figures.push(
        countFigure(summary.parsed, {
            key: 'parsedRows',
            label: 'Rows parsed',
            group: 'input',
            basis:
                chequeRows > 0
                    ? `${summary.parsed.toLocaleString()} GL posting(s) + ${chequeRows.toLocaleString()} cheque register row(s) normalised from ${summary.dataRows.toLocaleString()} data rows (blank padding skipped).`
                    : `${summary.parsed.toLocaleString()} of ${summary.dataRows.toLocaleString()} data rows normalised into GL postings (blank padding rows skipped).`,
            assessment:
                failed === 0
                    ? 'Every data row parsed cleanly — the population the reconciliation runs on is complete.'
                    : `${failed.toLocaleString()} row(s) did not parse; they are excluded from every figure below and listed under parse errors — review before relying on the balance.`,
        })
    );

    const creditFils = sheetBalances.filter((s) => s.role !== 'register').reduce((a, s) => a + s.creditFils, 0);
    const debitFils = sheetBalances.filter((s) => s.role !== 'register').reduce((a, s) => a + s.debitFils, 0);
    figures.push(
        moneyFigure(-creditFils, {
            key: 'sumCredits',
            label: 'Σ credit postings',
            group: 'input',
            basis: 'Sum of every credit posting across the ledger sheet(s) — cheque issuances into the GL.',
            assessment: 'The money that flowed INTO the payable account (cheques written). Shown negative in engine sign; a liability grows on the credit side.',
        }),
        moneyFigure(debitFils, {
            key: 'sumDebits',
            label: 'Σ debit postings',
            group: 'input',
            basis: 'Sum of every debit posting across the ledger sheet(s) — cheque encashments out of the GL.',
            assessment: 'The money that flowed OUT of the payable account (cheques paid). Debits reduce the liability.',
        }),
        moneyFigure(summary.netFils, {
            key: 'netMovement',
            label: 'Net movement (Σ signed postings)',
            group: 'input',
            basis: 'Σ debits − Σ credits over all parsed postings.',
            inputs: ['sumCredits', 'sumDebits'],
            assessment:
                'The net change the ledger extract itself accounts for. It should reconstruct the balance movement over the window; a mismatch to the stated balance is the extract gap.',
        })
    );

    // ── Per-sheet balances ────────────────────────────────────────────────────
    for (const sb of sheetBalances) {
        if (sb.role === 'register') {
            figures.push(
                moneyFigure(sb.chequeFils ?? 0, {
                    key: `sheet:${sb.sheet}`,
                    label: `${sb.sheet} — Σ cheque amounts`,
                    group: 'sheet',
                    sheet: sb.sheet,
                    basis: sb.basis,
                    assessment: `The instrument register on "${sb.sheet}": ${(sb.chequeCount ?? 0).toLocaleString()} cheque(s) the ledger is matched against. Not a debit/credit balance — it is the source of truth for cheque identity.`,
                })
            );
        } else if (sb.role === 'skipped') {
            figures.push(
                countFigure(0, {
                    key: `sheet:${sb.sheet}`,
                    label: `${sb.sheet} — skipped`,
                    group: 'sheet',
                    sheet: sb.sheet,
                    basis: sb.basis,
                    assessment: `"${sb.sheet}" matched no known schema and contributes nothing to the balance.`,
                })
            );
        } else {
            figures.push(
                moneyFigure(sb.netFils, {
                    key: `sheet:${sb.sheet}`,
                    label: `${sb.sheet} — net balance`,
                    group: 'sheet',
                    sheet: sb.sheet,
                    basis: sb.basis,
                    assessment:
                        sb.statedEodFils !== undefined
                            ? `Sheet "${sb.sheet}" nets to ${money(sb.netFils)} from ${sb.parsedRows.toLocaleString()} posting(s); it also states a closing EoD balance of ${money(sb.statedEodFils)} (the ledger's own figure).`
                            : `Sheet "${sb.sheet}" nets to ${money(sb.netFils)} from ${sb.parsedRows.toLocaleString()} posting(s).`,
                })
            );
        }
    }

    if (input.mode === 'register' && input.reconciliation?.byBranch[0]) {
        explainRegister(figures, input.reconciliation.byBranch[0], input);
    } else if (input.mode === 'statement' && input.reconciliation?.byBranch[0]) {
        explainStatement(figures, input.reconciliation.byBranch[0]);
    } else {
        explainBreakdown(figures, input);
    }

    return figures;
}

/** GOAL-8: a running-balance statement run — no matching, no outstanding; the story
 *  is derived vs stated and the tie-out gap between them. */
function explainStatement(figures: ExplainedFigure[], block: NonNullable<Reconciliation['byBranch'][number]>): void {
    figures.push(
        moneyFigure(block.derivedBalanceFils ?? 0, {
            key: 'derivedBalance',
            label: 'Derived balance (from postings)',
            group: 'balance',
            basis: 'Σ signed postings across the statement (debit +, credit −), independent of the stated figure.',
            inputs: ['netMovement'],
            assessment: 'What the ledger movements themselves reconstruct — the number that must equal the stated closing balance.',
        }),
        moneyFigure(block.glBalanceFils, {
            key: 'glBalance',
            label: 'Stated End-Date EoD balance',
            group: 'balance',
            basis: "The End Date EoD Balance on the statement's final-day rows, negated into engine sign — the ledger's own closing figure.",
            assessment: 'The control total for a running-balance GL: the derived balance must reproduce it exactly.',
        })
    );
    if (block.extractGapFils !== undefined) {
        const gap = block.extractGapFils;
        figures.push(
            moneyFigure(gap, {
                key: 'extractGap',
                label: 'Tie-out gap (derived − stated)',
                group: 'reconciliation',
                basis: 'Derived balance − stated End-Date EoD balance.',
                inputs: ['derivedBalance', 'glBalance'],
                flag: gap !== 0,
                assessment: gap === 0
                    ? 'The postings reproduce the stated balance to the fil — the running-balance ledger ties out.'
                    : `${money(Math.abs(gap))} of movement is unaccounted for between the postings and the stated closing balance — investigate before relying on the VAT figure.`,
            })
        );
    }
}

function explainRegister(
    figures: ExplainedFigure[],
    block: NonNullable<Reconciliation['byBranch'][number]>,
    input: ExplainInput
): void {
    const stated = block.statedBalanceFils;
    const glBalance = block.glBalanceFils;

    figures.push(
        moneyFigure(glBalance, {
            key: 'glBalance',
            label: 'GL closing balance',
            group: 'balance',
            basis:
                stated !== undefined
                    ? "The stated End Date EoD Balance on the ledger's final-day rows — the bank's own authoritative closing figure for the GL."
                    : 'Derived from the postings (the file states no closing balance) — Σ signed postings as at the review date.',
            assessment:
                'The control total the reconciliation must tie to: outstanding cheques plus classified exceptions should reconstruct exactly this number.',
        })
    );

    if (block.derivedBalanceFils !== undefined) {
        figures.push(
            moneyFigure(block.derivedBalanceFils, {
                key: 'derivedBalance',
                label: 'Derived balance (from postings)',
                group: 'balance',
                basis: 'Σ signed postings across the ledger sheets, computed independently of the stated figure.',
                inputs: ['netMovement'],
                assessment:
                    'A cross-check on the stated balance. When the two agree the extract is complete; when they differ the ledger extract is missing movements (the extract gap).',
            })
        );
    }

    if (block.extractGapFils !== undefined) {
        const gap = block.extractGapFils;
        figures.push(
            moneyFigure(gap, {
                key: 'extractGap',
                label: 'Ledger extract gap',
                group: 'balance',
                basis: 'Derived balance − stated EoD balance.',
                inputs: ['derivedBalance', 'glBalance'],
                flag: gap !== 0,
                assessment:
                    gap === 0
                        ? 'The postings fully reconstruct the stated balance — the extract is complete.'
                        : `The postings do NOT fully reconstruct the stated balance: ${money(Math.abs(gap))} of movement is either missing from the extract or outside the EoD scope. Obtain a complete extract before treating the tie-out as final.`,
            })
        );
    }

    figures.push(
        moneyFigure(block.outstandingNetFils, {
            key: 'outstandingNet',
            label: 'Σ outstanding (all unmatched)',
            group: 'matching',
            basis: 'Σ signed over every unmatched posting after two-legged GL↔register matching — equals the derived balance by the tested invariant.',
            assessment: 'The full unmatched mass. It splits into outstanding cheques (the statement) and classified exceptions.',
        }),
        countFigure(block.oldCount, {
            key: 'oldItems',
            label: 'Old outstanding cheques (> 1 year)',
            group: 'reconciliation',
            basis: `${block.oldCount.toLocaleString()} outstanding cheque(s) issued more than a year before the review date, Σ ${money(block.oldFils)}.`,
            assessment: 'Aged manager’s cheques never encashed — the "old items" population that needs periodic write-back review.',
        }),
        countFigure(block.currentCount, {
            key: 'currentItems',
            label: 'Current outstanding cheques (< 1 year)',
            group: 'reconciliation',
            basis: `${block.currentCount.toLocaleString()} outstanding cheque(s) issued within the last year, Σ ${money(block.currentFils)}.`,
            assessment: 'Recently issued cheques still legitimately in circulation.',
        }),
        moneyFigure(block.oldFils + block.currentFils, {
            key: 'statementTotal',
            label: 'Statement total (outstanding cheques)',
            group: 'reconciliation',
            basis: 'Σ |amount| of outstanding cheques = old items + current items.',
            inputs: ['oldItems', 'currentItems'],
            assessment: 'What the outstanding-cheque statement adds up to — the part of the GL balance explained by live instruments.',
        })
    );

    if (block.classifiedFils !== undefined) {
        figures.push(
            moneyFigure(block.classifiedFils, {
                key: 'classified',
                label: 'Classified exceptions',
                group: 'exceptions',
                basis: 'Σ outstanding − statement total: unmatched mass that is NOT an outstanding cheque (non-issuance credits, batch/unmatched debits, register lag).',
                inputs: ['outstandingNet', 'statementTotal'],
                assessment: 'Movement explained by a named reason rather than a live cheque. Each class is listed on the Mismatched sheet with the action it needs.',
            })
        );
    }

    figures.push(
        moneyFigure(block.differenceFils, {
            key: 'difference',
            label: 'Difference (GL − outstanding)',
            group: 'reconciliation',
            basis: 'GL closing balance − Σ outstanding.',
            inputs: ['glBalance', 'outstandingNet'],
            assessment: 'Zero (within tolerance) means the GL is fully explained by the unmatched population.',
        }),
        moneyFigure(block.residualFils ?? block.differenceFils, {
            key: 'residual',
            label: 'Unexplained residual',
            group: 'reconciliation',
            flag: !block.balanced,
            basis: 'Difference after removing classified exceptions — what no cheque and no named exception accounts for.',
            inputs: ['difference', 'classified'],
            assessment: block.balanced
                ? 'Within tolerance — the reconciliation ties out; nothing left to chase.'
                : `${money(Math.abs(block.residualFils ?? block.differenceFils))} is unexplained. This is the true exception a reviewer must investigate before signing off.`,
        })
    );

    if (input.chequeCount !== undefined) {
        figures.push(
            countFigure(input.chequeCount, {
                key: 'chequeCount',
                label: 'Cheque register rows',
                group: 'matching',
                basis: `${input.chequeCount.toLocaleString()} instrument row(s) in the register${statesText(input.chequesByState)}.`,
                assessment: 'The instrument population matched against the ledger; each cheque resolves to a state (paid / outstanding / stopped / pre-window).',
            })
        );
    }
}

function statesText(byState?: Partial<Record<ChequeState, number>>): string {
    if (!byState) {
        return '';
    }
    const parts = Object.entries(byState)
        .filter(([, n]) => n)
        .map(([s, n]) => `${n!.toLocaleString()} ${s.toLowerCase().replace(/_/g, ' ')}`);
    return parts.length ? ` (${parts.join(', ')})` : '';
}

function explainBreakdown(figures: ExplainedFigure[], input: ExplainInput): void {
    const glTotalFils = input.balances.reduce((a, b) => a + b.balanceFils, 0);
    figures.push(
        moneyFigure(glTotalFils, {
            key: 'glBalanceTotal',
            label: 'GL balance (all branches)',
            group: 'balance',
            basis: `Σ signed postings per (entity, GL, branch) as at ${input.asOf ?? 'the latest post date'}, summed over ${input.balances.length.toLocaleString()} branch balance(s).`,
            assessment: 'The derived control total the outstanding items must reconcile to, branch by branch.',
        })
    );

    if (input.matching) {
        const m = input.matching;
        figures.push(
            moneyFigure(m.matchedFils, {
                key: 'matched',
                label: 'Matched (cleared) amount',
                group: 'matching',
                basis: `Σ |amount| paired debit↔credit by FIFO offset within the match key [${m.matchKey.join(', ')}], across ${m.matchedSetCount.toLocaleString()} cleared set(s).`,
                assessment: 'Postings that cancelled out — cheques issued and encashed within the window. They leave nothing outstanding.',
            }),
            moneyFigure(m.netOutstandingFils, {
                key: 'outstandingNet',
                label: 'Σ outstanding (net)',
                group: 'matching',
                basis: 'Σ signed over every leftover after offsetting — equals Σ signed postings by the tie-out invariant.',
                assessment: 'The unmatched remainder that must equal the GL balance for the account to reconcile.',
            })
        );
    }

    if (input.reconciliation) {
        figures.push(
            moneyFigure(input.reconciliation.totalAbsDifferenceFils, {
                key: 'totalDifference',
                label: 'Total absolute difference',
                group: 'reconciliation',
                basis: 'Σ |GL balance − Σ outstanding| across all branches.',
                inputs: ['glBalanceTotal', 'outstandingNet'],
                flag: !input.reconciliation.balanced,
                assessment: input.reconciliation.balanced
                    ? 'Zero within tolerance — every branch reconciles.'
                    : 'Non-zero — at least one branch does not reconcile; the difference is the amount to investigate.',
            })
        );
    }

    if (input.exceptionsSummary) {
        figures.push(
            countFigure(input.exceptionsSummary.total, {
                key: 'exceptions',
                label: 'Reconciling exceptions',
                group: 'exceptions',
                basis: `Every outstanding item classified into a reviewer-facing reason (${Object.entries(input.exceptionsSummary.byReason)
                    .map(([r, n]) => `${n} ${r.toLowerCase().replace(/_/g, ' ')}`)
                    .join(', ')}).`,
                assessment: 'The full worklist: each exception is one leftover posting with the reason it did not clear and the action it needs.',
            })
        );
    }
}

/** filsToBhd re-export convenience for callers that want the decimal alongside the display. */
export { filsToBhd };
